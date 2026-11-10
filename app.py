from __future__ import annotations

import hmac
import json
import logging
import os
import shutil
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Any
from uuid import UUID, uuid4

from flask import Flask, Response, jsonify, render_template, request, send_file
from flask_sock import Sock

from processor.services.excel_processor import ProcessingError
from processor.services.history_store import (
    DATABASE_PATH as HISTORY_DATABASE_PATH,
    RETENTION_SECONDS,
    add_history_entry,
    delete_history_entry,
    get_history_entry,
    initialize_history_store,
    list_history,
    prune_expired_history,
    update_history_entry,
)
from processor.services.new_format_processor import (
    NewBankStatementProcessor,
    ProcessingCancelled,
)


ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xls"}
RESULT_TTL_SECONDS = int(os.getenv("RESULT_TTL_MINUTES", "30")) * 60
HISTORY_FILES_DIR = Path(
    os.getenv("HISTORY_FILES_DIR", str(HISTORY_DATABASE_PATH.parent / "history_files"))
).expanduser()
MAX_WS_MANIFEST_BYTES = 64 * 1024
MAX_WS_FILES = 100
PAYMENTS_API_TOKEN = os.getenv("PAYMENTS_API_TOKEN", "").strip()
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "100")) * 1024 * 1024
app.config["SOCK_SERVER_OPTIONS"] = {
    "max_message_size": app.config["MAX_CONTENT_LENGTH"],
    "ping_interval": 25,
}
sock = Sock(app)
initialize_history_store()


@dataclass
class Operation:
    operation_id: str
    client_id: str
    statement_files: list[str] = field(default_factory=list, repr=False)
    reference_file: str = field(default="", repr=False)
    detail_files: list[str] = field(default_factory=list, repr=False)
    created_at: float = field(default_factory=time.time)
    started_at: float | None = None
    finished_at: float | None = None
    status: str = "queued"
    progress: int = 0
    message: str = "Операция поставлена в очередь"
    error: str = ""
    result_bytes: bytes | None = None
    stats: dict[str, int] = field(default_factory=dict)
    version: int = 0
    cancel_event: threading.Event = field(default_factory=threading.Event, repr=False)
    condition: threading.Condition = field(default_factory=threading.Condition, repr=False)

    def snapshot(self) -> dict[str, Any]:
        with self.condition:
            return {
                "operation_id": self.operation_id,
                "status": self.status,
                "progress": self.progress,
                "message": self.message,
                "error": self.error,
                "created_at": self.created_at,
                "started_at": self.started_at,
                "finished_at": self.finished_at,
                "result_available": self.result_bytes is not None,
                "stats": dict(self.stats),
                "version": self.version,
            }


operations: dict[str, Operation] = {}
operations_lock = threading.RLock()
executor = ThreadPoolExecutor(
    max_workers=int(os.getenv("PROCESSING_WORKERS", "4")),
    thread_name_prefix="excel-operation",
)


@app.get("/")
def index() -> str:
    return render_template("index.html")


@app.get("/history")
def history_page() -> Response:
    response = app.make_response(render_template("history.html"))
    response.headers["Cache-Control"] = "private, no-store"
    response.headers["X-Robots-Tag"] = "noindex, nofollow"
    return response


@app.get("/api/history")
def operation_history() -> Response | tuple[Response, int]:
    items = list_history()
    for item in items:
        item["downloads"] = _history_downloads(item)
    response = jsonify({"items": items, "retention_days": 30})
    response.headers["Cache-Control"] = "private, no-store"
    return response


@app.get("/history/<operation_id>/input/<file_group>/<int:file_index>")
def download_history_input(
    operation_id: str,
    file_group: str,
    file_index: int,
) -> Response | tuple[Response, int]:
    entry = _history_entry(operation_id)
    if entry is None:
        return _error("Файл в истории не найден.", 404)

    if file_group == "statements":
        filenames = entry["statement_files"]
    elif file_group == "reference":
        filenames = [entry["reference_file"]]
    elif file_group == "details":
        filenames = entry["detail_files"]
    else:
        return _error("Файл в истории не найден.", 404)

    if file_index < 0 or file_index >= len(filenames):
        return _error("Файл в истории не найден.", 404)
    filename = filenames[file_index]
    path = _stored_history_file(entry["operation_id"], file_group, filename)
    if path is None or not path.is_file():
        return _error("Файл истории больше недоступен.", 404)
    return _history_attachment(path, filename)


@app.get("/history/<operation_id>/result")
def download_history_result(operation_id: str) -> Response | tuple[Response, int]:
    entry = _history_entry(operation_id)
    if entry is None or entry["status"] != "completed":
        return _error("Результат обработки не найден.", 404)
    path = _stored_history_file(entry["operation_id"], "output", "registry.xlsx")
    if path is None or not path.is_file():
        return _error("Результат обработки недоступен.", 404)
    return _history_attachment(path, "registry.xlsx")


@app.post("/operations")
def create_operation() -> tuple[Response, int]:
    client_id = _normalized_uuid(request.form.get("client_id", ""))
    if not client_id:
        return _error("Некорректный идентификатор браузера.")

    statement_files = _uploaded_files("statement_files")
    detail_files = _uploaded_files("detail_files")
    reference_file = request.files.get("reference_file")
    if not statement_files:
        return _error("Выберите хотя бы одну банковскую выписку.")
    if reference_file is None or not reference_file.filename:
        return _error("Выберите справочник заемщиков.")

    all_files = [*statement_files, reference_file, *detail_files]
    invalid_names = [item.filename for item in all_files if not _is_excel(item.filename)]
    if invalid_names:
        return _error(
            "Поддерживаются только файлы .xlsx, .xlsm, .xltx и .xls: "
            + ", ".join(invalid_names)
        )

    operation = Operation(operation_id=str(uuid4()), client_id=client_id)
    temp_dir = Path(tempfile.mkdtemp(prefix=f"payment-registry-{operation.operation_id}-"))
    history_dir = _history_operation_dir(operation.operation_id)
    try:
        history_dir.mkdir(parents=True, exist_ok=False)
        statement_paths = _save_uploads(statement_files, history_dir / "statements")
        detail_paths = _save_uploads(detail_files, history_dir / "details")
        reference_path = _save_uploads([reference_file], history_dir / "reference")[0]
        operation.statement_files = [path.name for path in statement_paths]
        operation.reference_file = reference_path.name
        operation.detail_files = [path.name for path in detail_paths]
        os.utime(history_dir, (operation.created_at, operation.created_at))
        add_history_entry(
            operation_id=operation.operation_id,
            client_id=operation.client_id,
            created_at=operation.created_at,
            statement_files=operation.statement_files,
            reference_file=operation.reference_file,
            detail_files=operation.detail_files,
        )
        with operations_lock:
            operations[operation.operation_id] = operation
        executor.submit(
            _run_operation,
            operation,
            statement_paths,
            reference_path,
            detail_paths,
            temp_dir,
        )
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        shutil.rmtree(history_dir, ignore_errors=True)
        with operations_lock:
            operations.pop(operation.operation_id, None)
        try:
            delete_history_entry(operation.operation_id, operation.client_id)
        except Exception:
            logger.exception("Could not remove incomplete history entry")
        return _error("Не удалось принять файлы для обработки.", 500)

    return jsonify(operation.snapshot()), 202


@app.get("/operations/<operation_id>")
def operation_status(operation_id: str) -> Response | tuple[Response, int]:
    operation = _owned_operation(operation_id, request.args.get("client_id", ""))
    if operation is None:
        return _error("Операция не найдена.", 404)
    return jsonify(operation.snapshot())


@app.post("/operations/<operation_id>/cancel")
def cancel_operation(operation_id: str) -> Response | tuple[Response, int]:
    payload = request.get_json(silent=True) or {}
    operation = _owned_operation(operation_id, payload.get("client_id", ""))
    if operation is None:
        return _error("Операция не найдена.", 404)

    history_changed = False
    with operation.condition:
        if operation.status in {"queued", "processing", "cancelling"}:
            operation.cancel_event.set()
            history_changed = operation.status != "cancelling"
            operation.status = "cancelling"
            operation.message = "Отмена операции..."
            operation.version += 1
            operation.condition.notify_all()
    if history_changed:
        _sync_operation_history(operation)
    return jsonify(operation.snapshot())


@app.get("/operations/<operation_id>/result")
def operation_result(operation_id: str) -> Response | tuple[Response, int]:
    operation = _owned_operation(operation_id, request.args.get("client_id", ""))
    if operation is None:
        return _error("Операция не найдена.", 404)

    with operation.condition:
        result_bytes = operation.result_bytes
        status = operation.status
    if status != "completed" or result_bytes is None:
        return _error("Результат еще не готов.", 409)

    response = send_file(
        BytesIO(result_bytes),
        as_attachment=True,
        download_name="registry.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response.headers["Cache-Control"] = "private, no-store"
    return response


@sock.route("/ws")
def operation_socket(ws: Any) -> None:
    operation = _owned_operation(
        request.args.get("operation_id", ""),
        request.args.get("client_id", ""),
    )
    if operation is None:
        ws.send(json.dumps({"status": "error", "error": "Операция не найдена."}))
        return

    last_version = -1
    try:
        while True:
            with operation.condition:
                operation.condition.wait_for(
                    lambda: operation.version != last_version,
                    timeout=15,
                )
                snapshot = operation.snapshot()
                last_version = operation.version
            ws.send(json.dumps(snapshot, ensure_ascii=False))
            if snapshot["status"] in {"completed", "cancelled", "error"}:
                return
    except Exception:
        return


@sock.route("/api/v1/ws/process")
def process_websocket_api(ws: Any) -> None:
    operation_id = str(uuid4())
    if not PAYMENTS_API_TOKEN:
        _send_websocket_error(
            ws,
            operation_id,
            "service_unconfigured",
            "PAYMENTS_API_TOKEN is not configured.",
        )
        return
    authorization = request.headers.get("Authorization", "").encode("utf-8")
    expected_authorization = f"Bearer {PAYMENTS_API_TOKEN}".encode("utf-8")
    if not hmac.compare_digest(authorization, expected_authorization):
        _send_websocket_error(ws, operation_id, "unauthorized", "A valid bearer token is required.")
        return

    temp_dir = Path(tempfile.mkdtemp(prefix=f"payment-registry-ws-{operation_id}-"))
    request_id: str | None = None
    actor: dict[str, str] | None = None
    try:
        first_message = ws.receive(timeout=30)
        if not isinstance(first_message, str):
            raise ValueError("The first WebSocket message must be a JSON text manifest.")
        if len(first_message.encode("utf-8")) > MAX_WS_MANIFEST_BYTES:
            raise ValueError("The WebSocket manifest is too large.")
        try:
            manifest = json.loads(first_message)
        except json.JSONDecodeError as exc:
            raise ValueError("The WebSocket manifest is not valid JSON.") from exc

        request_id, actor = _parse_websocket_context(manifest)
        if manifest.get("type") != "start":
            raise ValueError('The first message must have type "start".')
        file_specs = _parse_websocket_manifest(manifest)
        audit_context = {"request_id": request_id, "actor": actor}
        total_bytes = sum(item["size"] for item in file_specs)
        _send_websocket_json(
            ws,
            {
                "type": "accepted",
                "operation_id": operation_id,
                "status": "receiving",
                "total_files": len(file_specs),
                "total_bytes": total_bytes,
                **audit_context,
            },
        )

        paths: dict[str, list[Path]] = {"statements": [], "reference": [], "details": []}
        received_bytes = 0
        for index, item in enumerate(file_specs, start=1):
            file_data = ws.receive()
            if not isinstance(file_data, bytes):
                raise ValueError(f"File {index} must be sent as a binary WebSocket message.")
            if len(file_data) != item["size"]:
                raise ValueError(
                    f"File {index} size mismatch: expected {item['size']} bytes, got {len(file_data)}."
                )

            destination = _unique_upload_path(
                temp_dir / item["group"],
                item["name"],
            )
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(file_data)
            paths[item["group"]].append(destination)
            received_bytes += len(file_data)
            _send_websocket_json(
                ws,
                {
                    "type": "upload_progress",
                    "operation_id": operation_id,
                    "received_files": index,
                    "total_files": len(file_specs),
                    "received_bytes": received_bytes,
                    "total_bytes": total_bytes,
                    **audit_context,
                },
            )

        _send_websocket_json(
            ws,
            {
                "type": "progress",
                "operation_id": operation_id,
                "status": "processing",
                "progress": 1,
                "message": "Processing started.",
                **audit_context,
            },
        )
        result = NewBankStatementProcessor().process_many(
            statement_paths=paths["statements"],
            reference_path=paths["reference"][0],
            output_dir=temp_dir / "output",
            detail_paths=paths["details"],
            on_progress=lambda progress, message: _send_websocket_json(
                ws,
                {
                    "type": "progress",
                    "operation_id": operation_id,
                    "status": "processing",
                    "progress": progress,
                    "message": message,
                    **audit_context,
                },
            ),
        )
        result_bytes = result.registry_path.read_bytes()
        stats = {
            "total": result.total_statement_rows,
            "generated": result.generated_records,
            "matched": result.matched_rows,
            "notFound": result.not_found_rows,
            "excluded": result.excluded_rows,
            "split": result.split_rows,
        }
        _send_websocket_json(
            ws,
            {
                "type": "complete",
                "operation_id": operation_id,
                "status": "completed",
                "progress": 100,
                "filename": "registry.xlsx",
                "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "size": len(result_bytes),
                "stats": stats,
                **audit_context,
            },
        )
        ws.send(result_bytes)
    except ValueError as exc:
        _send_websocket_error(ws, operation_id, "invalid_request", str(exc), request_id, actor)
    except ProcessingError as exc:
        _send_websocket_error(ws, operation_id, "processing_error", str(exc), request_id, actor)
    except Exception:
        _send_websocket_error(
            ws,
            operation_id,
            "internal_error",
            "Could not process the uploaded Excel files.",
            request_id,
            actor,
        )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


@app.errorhandler(413)
def upload_too_large(_error: Exception) -> tuple[Response, int]:
    limit = os.getenv("MAX_UPLOAD_MB", "100")
    return jsonify(error=f"Общий размер загрузки превышает {limit} МБ."), 413


def _run_operation(
    operation: Operation,
    statement_paths: list[Path],
    reference_path: Path,
    detail_paths: list[Path],
    temp_dir: Path,
) -> None:
    try:
        _update_operation(operation, status="processing", progress=1, message="Начинаем обработку")
        result = NewBankStatementProcessor().process_many(
            statement_paths=statement_paths,
            reference_path=reference_path,
            output_dir=temp_dir / "output",
            detail_paths=detail_paths,
            should_cancel=operation.cancel_event.is_set,
            on_progress=lambda progress, message: _update_operation(
                operation,
                status="processing",
                progress=progress,
                message=message,
            ),
        )
        if operation.cancel_event.is_set():
            raise ProcessingCancelled("Операция отменена")

        result_bytes = result.registry_path.read_bytes()
        result_archive = _history_operation_dir(operation.operation_id) / "output" / "registry.xlsx"
        result_archive.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(result.registry_path, result_archive)
        os.utime(
            _history_operation_dir(operation.operation_id),
            (operation.created_at, operation.created_at),
        )
        stats = {
            "total": result.total_statement_rows,
            "generated": result.generated_records,
            "matched": result.matched_rows,
            "notFound": result.not_found_rows,
            "excluded": result.excluded_rows,
            "split": result.split_rows,
        }
        _update_operation(
            operation,
            status="completed",
            progress=100,
            message="Реестр сформирован",
            result_bytes=result_bytes,
            stats=stats,
            finished=True,
        )
    except ProcessingCancelled:
        _update_operation(
            operation,
            status="cancelled",
            message="Операция отменена",
            result_bytes=None,
            finished=True,
        )
    except ProcessingError as exc:
        _update_operation(
            operation,
            status="error",
            message="Обработка завершилась с ошибкой",
            error=str(exc),
            finished=True,
        )
    except Exception:
        _update_operation(
            operation,
            status="error",
            message="Обработка завершилась с ошибкой",
            error="Не удалось обработать файлы. Проверьте структуру Excel-файлов.",
            finished=True,
        )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _update_operation(
    operation: Operation,
    *,
    status: str | None = None,
    progress: int | None = None,
    message: str | None = None,
    error: str | None = None,
    result_bytes: bytes | None = None,
    stats: dict[str, int] | None = None,
    finished: bool = False,
) -> None:
    history_changed = False
    with operation.condition:
        previous_status = operation.status
        if operation.cancel_event.is_set() and status == "processing":
            status = "cancelling"
            message = "Отмена операции..."
        if operation.started_at is None and status in {"processing", "cancelling"}:
            operation.started_at = time.time()
        if status is not None:
            operation.status = status
        if progress is not None:
            operation.progress = max(0, min(100, int(progress)))
        if message is not None:
            operation.message = message
        if error is not None:
            operation.error = error
        if result_bytes is not None:
            operation.result_bytes = result_bytes
        if stats is not None:
            operation.stats = stats
        if finished:
            operation.finished_at = time.time()
        history_changed = operation.status != previous_status or finished
        operation.version += 1
        operation.condition.notify_all()
    if history_changed:
        _sync_operation_history(operation)


def _sync_operation_history(operation: Operation) -> None:
    with operation.condition:
        history = {
            "operation_id": operation.operation_id,
            "client_id": operation.client_id,
            "status": operation.status,
            "started_at": operation.started_at,
            "finished_at": operation.finished_at,
            "stats": dict(operation.stats),
            "error": operation.error,
        }
    try:
        update_history_entry(**history)
    except Exception:
        logger.exception("Could not update history for operation %s", operation.operation_id)


def _owned_operation(operation_id: str, client_id: str) -> Operation | None:
    normalized_operation_id = _normalized_uuid(operation_id)
    normalized_client_id = _normalized_uuid(client_id)
    if not normalized_operation_id or not normalized_client_id:
        return None
    with operations_lock:
        operation = operations.get(normalized_operation_id)
    if operation is None or operation.client_id != normalized_client_id:
        return None
    return operation


def _history_entry(operation_id: str) -> dict | None:
    normalized_operation_id = _normalized_uuid(operation_id)
    if not normalized_operation_id:
        return None
    return get_history_entry(normalized_operation_id)


def _history_operation_dir(operation_id: str) -> Path:
    return HISTORY_FILES_DIR / operation_id


def _stored_history_file(operation_id: str, file_group: str, filename: str) -> Path | None:
    if not isinstance(filename, str) or not filename:
        return None
    safe_name = _safe_upload_name(filename)
    if safe_name != filename:
        return None
    group_dir = (_history_operation_dir(operation_id) / file_group).resolve()
    path = (group_dir / safe_name).resolve()
    if path.parent != group_dir:
        return None
    return path


def _stored_history_file_exists(operation_id: str, file_group: str, filename: str) -> bool:
    path = _stored_history_file(operation_id, file_group, filename)
    return path is not None and path.is_file()


def _history_attachment(path: Path, filename: str) -> Response:
    response = send_file(path, as_attachment=True, download_name=filename)
    response.headers["Cache-Control"] = "private, no-store"
    return response


def _history_downloads(entry: dict) -> dict[str, Any]:
    statements = entry["statement_files"]
    details = entry["detail_files"]
    return {
        "statements": [
            _stored_history_file_exists(entry["operation_id"], "statements", name)
            for name in statements
        ],
        "reference": [
            _stored_history_file_exists(
                entry["operation_id"], "reference", entry["reference_file"]
            )
        ],
        "details": [
            _stored_history_file_exists(entry["operation_id"], "details", name)
            for name in details
        ],
        "result": bool(
            entry["status"] == "completed"
            and _stored_history_file_exists(entry["operation_id"], "output", "registry.xlsx")
        ),
    }


def _normalized_uuid(value: str) -> str:
    try:
        return str(UUID(str(value)))
    except (TypeError, ValueError, AttributeError):
        return ""


def _uploaded_files(field_name: str) -> list:
    return [item for item in request.files.getlist(field_name) if item.filename]


def _is_excel(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def _safe_upload_name(filename: str) -> str:
    name = Path(filename.replace("\\", "/")).name.strip()
    return name or "upload.xlsx"


def _save_uploads(uploaded_files: list, directory: Path) -> list[Path]:
    directory.mkdir(parents=True, exist_ok=True)
    paths: list[Path] = []
    for uploaded_file in uploaded_files:
        filename = _safe_upload_name(uploaded_file.filename)
        destination = directory / filename
        counter = 2
        while destination.exists():
            destination = directory / f"{Path(filename).stem} ({counter}){Path(filename).suffix}"
            counter += 1
        uploaded_file.save(destination)
        paths.append(destination)
    return paths


def _parse_websocket_context(manifest: Any) -> tuple[str, dict[str, str]]:
    if not isinstance(manifest, dict):
        raise ValueError("The WebSocket manifest must be a JSON object.")
    request_id_value = manifest.get("request_id")
    if isinstance(request_id_value, bool) or not isinstance(request_id_value, (str, int)):
        raise ValueError("A payments request_id string or integer is required.")
    request_id = str(request_id_value).strip()
    if not request_id or len(request_id) > 128:
        raise ValueError("The payments request_id must contain 1 to 128 characters.")

    actor_value = manifest.get("actor")
    if not isinstance(actor_value, dict):
        raise ValueError("The actor object with id and username is required.")
    actor_id_value = actor_value.get("id")
    actor_username = actor_value.get("username")
    if isinstance(actor_id_value, bool) or not isinstance(actor_id_value, (str, int)):
        raise ValueError("The actor id must be a string or integer.")
    actor_id = str(actor_id_value).strip()
    if not actor_id or not isinstance(actor_username, str) or not actor_username.strip():
        raise ValueError("The actor must include a non-empty id and username.")
    actor = {"id": actor_id, "username": actor_username.strip()[:150]}
    display_name = actor_value.get("display_name")
    if display_name is not None:
        if not isinstance(display_name, str):
            raise ValueError("The actor display_name must be a string.")
        actor["display_name"] = display_name.strip()[:200]

    return request_id, actor


def _parse_websocket_manifest(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    statements = manifest.get("statements")
    reference = manifest.get("reference")
    details = manifest.get("details", [])
    if not isinstance(statements, list) or not statements:
        raise ValueError("At least one statement file is required.")
    if not isinstance(reference, dict):
        raise ValueError("One reference file is required.")
    if not isinstance(details, list):
        raise ValueError('The "details" value must be a list.')

    groups = [
        ("statements", statements),
        ("reference", [reference]),
        ("details", details),
    ]
    file_count = sum(len(items) for _, items in groups)
    if file_count > MAX_WS_FILES:
        raise ValueError(f"No more than {MAX_WS_FILES} files can be sent in one operation.")

    file_specs: list[dict[str, Any]] = []
    total_bytes = 0
    for group, items in groups:
        for item in items:
            if not isinstance(item, dict):
                raise ValueError("Each file entry must be an object with name and size fields.")
            raw_name = item.get("name")
            size = item.get("size")
            if not isinstance(raw_name, str) or not raw_name.strip():
                raise ValueError("Each file must have a non-empty name.")
            if isinstance(size, bool) or not isinstance(size, int) or size <= 0:
                raise ValueError("Each file must have a positive integer size in bytes.")

            name = _safe_upload_name(raw_name)
            if not _is_excel(name):
                raise ValueError(
                    f"Unsupported file type for {name}. Use .xlsx, .xlsm, .xltx, or .xls."
                )
            total_bytes += size
            if total_bytes > app.config["MAX_CONTENT_LENGTH"]:
                limit_mb = app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024)
                raise ValueError(f"The total upload size exceeds the {limit_mb} MB limit.")
            file_specs.append({"group": group, "name": name, "size": size})

    return file_specs


def _unique_upload_path(directory: Path, filename: str) -> Path:
    destination = directory / filename
    counter = 2
    while destination.exists():
        destination = directory / f"{Path(filename).stem} ({counter}){Path(filename).suffix}"
        counter += 1
    return destination


def _send_websocket_json(ws: Any, payload: dict[str, Any]) -> None:
    ws.send(json.dumps(payload, ensure_ascii=False))


def _send_websocket_error(
    ws: Any,
    operation_id: str,
    code: str,
    message: str,
    request_id: str | None = None,
    actor: dict[str, str] | None = None,
) -> None:
    payload: dict[str, Any] = {
        "type": "error",
        "operation_id": operation_id,
        "status": "error",
        "code": code,
        "error": message,
    }
    if request_id is not None:
        payload["request_id"] = request_id
    if actor is not None:
        payload["actor"] = actor
    try:
        _send_websocket_json(ws, payload)
    except Exception:
        return


def _cleanup_expired_operations() -> None:
    while True:
        time.sleep(60)
        cutoff = time.time() - RESULT_TTL_SECONDS
        with operations_lock:
            expired_ids = [
                operation_id
                for operation_id, operation in operations.items()
                if operation.finished_at is not None and operation.finished_at < cutoff
            ]
            for operation_id in expired_ids:
                operations.pop(operation_id, None)
        try:
            prune_expired_history()
            _prune_expired_history_files()
        except Exception:
            logger.exception("Could not prune expired processing history")


def _remove_stale_temp_directories() -> None:
    temp_root = Path(tempfile.gettempdir())
    for directory in temp_root.glob("payment-registry-*"):
        if directory.is_dir():
            shutil.rmtree(directory, ignore_errors=True)


def _prune_expired_history_files() -> None:
    if not HISTORY_FILES_DIR.is_dir():
        return
    cutoff = time.time() - RETENTION_SECONDS
    with operations_lock:
        active_ids = set(operations)
    for directory in HISTORY_FILES_DIR.iterdir():
        if not directory.is_dir() or directory.name in active_ids:
            continue
        try:
            if directory.stat().st_mtime < cutoff:
                shutil.rmtree(directory, ignore_errors=True)
        except OSError:
            logger.exception("Could not prune history files in %s", directory)


def _error(message: str, status: int = 400) -> tuple[Response, int]:
    return jsonify(error=message), status


_remove_stale_temp_directories()
_prune_expired_history_files()
threading.Thread(
    target=_cleanup_expired_operations,
    name="operation-cleanup",
    daemon=True,
).start()


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", "8000")),
        debug=False,
        threaded=True,
    )

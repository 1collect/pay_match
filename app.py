from __future__ import annotations

import json
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
from processor.services.new_format_processor import (
    NewBankStatementProcessor,
    ProcessingCancelled,
)


ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xls"}
RESULT_TTL_SECONDS = int(os.getenv("RESULT_TTL_MINUTES", "30")) * 60

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_MB", "100")) * 1024 * 1024
sock = Sock(app)


@dataclass
class Operation:
    operation_id: str
    client_id: str
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
    try:
        statement_paths = _save_uploads(statement_files, temp_dir / "statements")
        detail_paths = _save_uploads(detail_files, temp_dir / "details")
        reference_path = _save_uploads([reference_file], temp_dir / "reference")[0]
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
        with operations_lock:
            operations.pop(operation.operation_id, None)
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

    with operation.condition:
        if operation.status in {"queued", "processing", "cancelling"}:
            operation.cancel_event.set()
            operation.status = "cancelling"
            operation.message = "Отмена операции..."
            operation.version += 1
            operation.condition.notify_all()
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
    with operation.condition:
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
        operation.version += 1
        operation.condition.notify_all()


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


def _remove_stale_temp_directories() -> None:
    temp_root = Path(tempfile.gettempdir())
    for directory in temp_root.glob("payment-registry-*"):
        if directory.is_dir():
            shutil.rmtree(directory, ignore_errors=True)


def _error(message: str, status: int = 400) -> tuple[Response, int]:
    return jsonify(error=message), status


_remove_stale_temp_directories()
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

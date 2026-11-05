from __future__ import annotations

import json
import os
import sqlite3
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator


RETENTION_SECONDS = 30 * 24 * 60 * 60
DEFAULT_DATABASE_PATH = Path(__file__).resolve().parents[2] / "history.sqlite3"
DATABASE_PATH = Path(os.getenv("HISTORY_DB_PATH", str(DEFAULT_DATABASE_PATH))).expanduser()


@contextmanager
def _connection() -> Iterator[sqlite3.Connection]:
    DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DATABASE_PATH, timeout=30)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA busy_timeout = 30000")
    try:
        yield connection
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def initialize_history_store() -> None:
    with _connection() as connection:
        connection.execute("PRAGMA journal_mode = WAL")
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS operation_history (
                operation_id TEXT PRIMARY KEY,
                client_id TEXT NOT NULL,
                created_at REAL NOT NULL,
                started_at REAL,
                finished_at REAL,
                status TEXT NOT NULL,
                statement_files TEXT NOT NULL,
                reference_file TEXT NOT NULL,
                detail_files TEXT NOT NULL,
                stats TEXT NOT NULL DEFAULT '{}',
                error TEXT NOT NULL DEFAULT ''
            )
            """
        )
        connection.execute(
            "CREATE INDEX IF NOT EXISTS idx_operation_history_client_created "
            "ON operation_history (client_id, created_at DESC)"
        )
        connection.execute(
            """
            UPDATE operation_history
            SET status = 'interrupted',
                finished_at = COALESCE(finished_at, ?),
                error = 'Операция прервана после перезапуска сервера.'
            WHERE status IN ('queued', 'processing', 'cancelling')
            """,
            (time.time(),),
        )
        _prune_expired(connection)


def add_history_entry(
    *,
    operation_id: str,
    client_id: str,
    created_at: float,
    statement_files: list[str],
    reference_file: str,
    detail_files: list[str],
) -> None:
    with _connection() as connection:
        connection.execute(
            """
            INSERT INTO operation_history (
                operation_id, client_id, created_at, status,
                statement_files, reference_file, detail_files
            ) VALUES (?, ?, ?, 'queued', ?, ?, ?)
            """,
            (
                operation_id,
                client_id,
                created_at,
                json.dumps(statement_files, ensure_ascii=False),
                reference_file,
                json.dumps(detail_files, ensure_ascii=False),
            ),
        )
        _prune_expired(connection)


def update_history_entry(
    *,
    operation_id: str,
    client_id: str,
    status: str,
    started_at: float | None,
    finished_at: float | None,
    stats: dict[str, int],
    error: str,
) -> None:
    with _connection() as connection:
        connection.execute(
            """
            UPDATE operation_history
            SET status = ?,
                started_at = COALESCE(?, started_at),
                finished_at = COALESCE(?, finished_at),
                stats = ?,
                error = ?
            WHERE operation_id = ? AND client_id = ?
            """,
            (
                status,
                started_at,
                finished_at,
                json.dumps(stats, ensure_ascii=False),
                error,
                operation_id,
                client_id,
            ),
        )
        _prune_expired(connection)


def delete_history_entry(operation_id: str, client_id: str) -> None:
    with _connection() as connection:
        connection.execute(
            "DELETE FROM operation_history WHERE operation_id = ? AND client_id = ?",
            (operation_id, client_id),
        )


def prune_expired_history() -> None:
    with _connection() as connection:
        _prune_expired(connection)


def list_history(client_id: str) -> list[dict]:
    with _connection() as connection:
        _prune_expired(connection)
        rows = connection.execute(
            """
            SELECT operation_id, created_at, started_at, finished_at, status,
                   statement_files, reference_file, detail_files, stats, error
            FROM operation_history
            WHERE client_id = ?
            ORDER BY created_at DESC
            """,
            (client_id,),
        ).fetchall()

    items = []
    for row in rows:
        items.append(
            {
                "operation_id": row["operation_id"],
                "created_at": row["created_at"],
                "started_at": row["started_at"],
                "finished_at": row["finished_at"],
                "status": row["status"],
                "statement_files": _load_json(row["statement_files"], []),
                "reference_file": row["reference_file"],
                "detail_files": _load_json(row["detail_files"], []),
                "stats": _load_json(row["stats"], {}),
                "error": row["error"],
            }
        )
    return items


def _prune_expired(connection: sqlite3.Connection) -> None:
    connection.execute(
        "DELETE FROM operation_history WHERE created_at < ?",
        (time.time() - RETENTION_SECONDS,),
    )


def _load_json(value: str, fallback):
    try:
        return json.loads(value)
    except (TypeError, json.JSONDecodeError):
        return fallback

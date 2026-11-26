import unittest
import time
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from uuid import uuid4
from zipfile import ZipFile

from app import Operation, app, operations, operations_lock
from processor.services import history_store


class WebTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = app.test_client()
        self.client_id = str(uuid4())

    def tearDown(self) -> None:
        with operations_lock:
            operations.clear()

    def test_index_is_available(self) -> None:
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn("Формирование реестра".encode(), response.data)

    def test_create_operation_requires_statement(self) -> None:
        response = self.client.post("/operations", data={"client_id": self.client_id})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            response.get_json(),
            {"error": "Выберите хотя бы одну банковскую выписку."},
        )

    def test_operation_is_visible_only_to_its_client(self) -> None:
        operation = Operation(operation_id=str(uuid4()), client_id=self.client_id)
        with operations_lock:
            operations[operation.operation_id] = operation

        owner_response = self.client.get(
            f"/operations/{operation.operation_id}?client_id={self.client_id}"
        )
        stranger_response = self.client.get(
            f"/operations/{operation.operation_id}?client_id={uuid4()}"
        )

        self.assertEqual(owner_response.status_code, 200)
        self.assertEqual(stranger_response.status_code, 404)

    def test_cancel_sets_operation_flag(self) -> None:
        operation = Operation(operation_id=str(uuid4()), client_id=self.client_id)
        with operations_lock:
            operations[operation.operation_id] = operation

        response = self.client.post(
            f"/operations/{operation.operation_id}/cancel",
            json={"client_id": self.client_id},
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["status"], "cancelling")
        self.assertTrue(operation.cancel_event.is_set())

    def test_history_api_returns_operations_from_all_clients(self) -> None:
        client_ids = [str(uuid4()), str(uuid4())]
        operation_ids = [str(uuid4()), str(uuid4())]

        with TemporaryDirectory() as temporary_directory:
            database_path = Path(temporary_directory) / "history.sqlite3"
            with patch.object(history_store, "DATABASE_PATH", database_path):
                history_store.initialize_history_store()
                for client_id, operation_id in zip(client_ids, operation_ids):
                    history_store.add_history_entry(
                        operation_id=operation_id,
                        client_id=client_id,
                        created_at=time.time(),
                        statement_files=["statement.xlsx"],
                        reference_file="reference.xlsx",
                        detail_files=[],
                    )

                response = self.client.get("/api/history")

        self.assertEqual(response.status_code, 200)
        self.assertCountEqual(
            [item["operation_id"] for item in response.get_json()["items"]],
            operation_ids,
        )

    def test_history_statements_zip_contains_only_operation_statements(self) -> None:
        operation_id = str(uuid4())
        filenames = ["Выписка ТОО.xlsx", "Выписка ТОО (1).xls"]
        entry = {"operation_id": operation_id, "statement_files": filenames}
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            statements = root / operation_id / "statements"
            statements.mkdir(parents=True)
            for index, filename in enumerate(filenames):
                (statements / filename).write_bytes(f"statement {index}".encode())
            (statements / "unlisted.xlsx").write_bytes(b"not in history")
            other = root / str(uuid4()) / "statements"
            other.mkdir(parents=True)
            (other / "other.xlsx").write_bytes(b"another operation")
            with patch("app.HISTORY_FILES_DIR", root), patch("app.get_history_entry", return_value=entry):
                response = self.client.get(f"/history/{operation_id}/statements.zip")
                self.assertEqual(response.status_code, 200)
                self.assertEqual(response.mimetype, "application/zip")
                self.assertIn(f"statements-{operation_id}.zip", response.headers["Content-Disposition"])
                self.assertEqual(response.headers["Cache-Control"], "private, no-store")
                with ZipFile(BytesIO(response.data)) as archive:
                    self.assertEqual(archive.namelist(), filenames)
                    for index, filename in enumerate(filenames):
                        self.assertEqual(archive.read(filename), f"statement {index}".encode())
                response.close()

    def test_history_statements_zip_rejects_missing_or_unsafe_files(self) -> None:
        operation_id = str(uuid4())
        with TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            statements = root / operation_id / "statements"
            statements.mkdir(parents=True)
            (statements / "available.xlsx").write_bytes(b"available")
            (root / operation_id / "outside.xlsx").write_bytes(b"outside")
            for filenames in [[], ["available.xlsx", "missing.xlsx"], ["../outside.xlsx"]]:
                with self.subTest(filenames=filenames), patch("app.HISTORY_FILES_DIR", root), patch(
                    "app.get_history_entry",
                    return_value={"operation_id": operation_id, "statement_files": filenames},
                ):
                    response = self.client.get(f"/history/{operation_id}/statements.zip")
                    self.assertEqual(response.status_code, 404)
                    self.assertEqual(response.mimetype, "application/json")

    def test_history_statements_zip_rejects_unknown_or_expired_operations(self) -> None:
        operation_id = str(uuid4())
        with TemporaryDirectory() as temporary_directory:
            with patch.object(history_store, "DATABASE_PATH", Path(temporary_directory) / "history.sqlite3"):
                history_store.initialize_history_store()
                history_store.add_history_entry(
                    operation_id=operation_id,
                    client_id=self.client_id,
                    created_at=time.time() - history_store.RETENTION_SECONDS - 1,
                    statement_files=["statement.xlsx"],
                    reference_file="reference.xlsx",
                    detail_files=[],
                )
                for value in [operation_id, str(uuid4()), "invalid-id"]:
                    with self.subTest(operation_id=value):
                        response = self.client.get(f"/history/{value}/statements.zip")
                        self.assertEqual(response.status_code, 404)


if __name__ == "__main__":
    unittest.main()

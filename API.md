# WebSocket API для `payments`

`/api/v1/ws/process` — серверный API для проекта `../payments`. `payments` создаёт собственную запись запуска и остаётся источником истории и авторства. Этот сервис проверяет API-токен, принимает файлы, обрабатывает их и возвращает `registry.xlsx`; он не сохраняет историю запусков.

## Настройка доступа

Задайте один и тот же секрет `PAYMENTS_API_TOKEN` в окружении обоих приложений. Подключение требует заголовок:

```http
Authorization: Bearer <PAYMENTS_API_TOKEN>
```

При пустом `PAYMENTS_API_TOKEN` endpoint отключён и возвращает событие `service_unconfigured`. Токен предназначен для серверного соединения `payments` → `pay_match`; не передавайте его браузеру.

## Протокол

Подключитесь к `ws://<host>/api/v1/ws/process` или `wss://<host>/api/v1/ws/process`. Первым сообщением отправьте JSON-манифест:

```json
{
  "type": "start",
  "request_id": "payments-run-1842",
  "actor": {
    "id": "42",
    "username": "operator",
    "display_name": "Имя сотрудника"
  },
  "statements": [
    {"name": "statement-1.xlsx", "size": 12345},
    {"name": "statement-2.xlsx", "size": 6789}
  ],
  "reference": {"name": "borrowers.xlsx", "size": 4321},
  "details": [
    {"name": "details.xlsx", "size": 9876}
  ]
}
```

`request_id` — ID записи запуска в `payments`; он возвращается во всех событиях, чтобы связать прогресс и результат с историей. `actor.id` и `actor.username` берите из аутентифицированного пользователя Django. `display_name` необязателен. `statements` должен содержать хотя бы одну выписку, `reference` — один справочник; `details` можно опустить или передать пустым массивом. Для каждого файла передайте имя и точный размер в байтах.

После события `accepted` передайте содержимое каждого файла отдельным бинарным WebSocket-сообщением в следующем порядке:

1. Выписки в порядке массива `statements`.
2. Справочник `reference`.
3. Расшифровки в порядке массива `details`.

JSON-события `upload_progress` и `progress` содержат `request_id`, `actor` и внутренний `operation_id`. При ошибке приходит `type: "error"` с кодом и сообщением. При успехе сервер отправляет JSON-событие `complete` со статистикой, а затем отдельным бинарным сообщением — готовый `registry.xlsx`.

Поддерживаются `.xlsx`, `.xlsm`, `.xltx` и `.xls`; общий размер файлов ограничен `MAX_UPLOAD_MB` (по умолчанию 100 МБ), количество — 100. Входной манифест ограничен 64 КБ.

## Пример серверного клиента на Python

Пример использует пакет `websockets` (`pip install websockets`). Функция `save_history_event` обозначает сохранение JSON-события в записи запуска на стороне `payments`.

```python
import json
import os
from pathlib import Path

from websockets.sync.client import connect


def file_info(path: Path) -> dict:
    return {"name": path.name, "size": path.stat().st_size}


def run_registry(*, run, user, statements, reference, details, output_path):
    files = [*statements, reference, *details]
    manifest = {
        "type": "start",
        "request_id": str(run.pk),
        "actor": {
            "id": str(user.pk),
            "username": user.get_username(),
            "display_name": user.get_full_name(),
        },
        "statements": [file_info(path) for path in statements],
        "reference": file_info(reference),
        "details": [file_info(path) for path in details],
    }

    with connect(
        os.environ["PAY_MATCH_WS_URL"],
        additional_headers={
            "Authorization": f"Bearer {os.environ['PAYMENTS_API_TOKEN']}"
        },
        max_size=None,
    ) as ws:
        ws.send(json.dumps(manifest))
        first_event = json.loads(ws.recv())
        if first_event.get("type") != "accepted":
            save_history_event(run, first_event)
            raise RuntimeError(first_event.get("error", "Registry request rejected"))

        save_history_event(run, first_event)
        for path in files:
            ws.send(path.read_bytes())

        while True:
            message = ws.recv()
            if isinstance(message, bytes):
                output_path.write_bytes(message)
                return output_path

            event = json.loads(message)
            save_history_event(run, event)
            if event.get("type") == "error":
                raise RuntimeError(event.get("error", "Registry processing failed"))
```

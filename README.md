# Notes API (FastAPI)

Простое API, возвращающее список заметок из базы данных SQLite.
Вся логика — в одном файле [`app.py`](app.py) (~40 строк). Схема БД и записи
создаются автоматически библиотекой [`sqlite-utils`](https://sqlite-utils.datasette.io).

## Файлы

| Файл | Назначение |
|------|------------|
| [`app.py`](app.py) | Всё приложение: база, данные и эндпоинт `GET /api/notes` |
| [`tests/test_api.py`](tests/test_api.py) | Автотесты (pytest) |
| [`postman_collection.json`](postman_collection.json) | Коллекция тестов Postman |
| [`requirements.txt`](requirements.txt) | Зависимости |

## База данных

Две таблицы:

- `users` — `id`, `login`;
- `notes` — `id`, `title`, `content`, `id_user` (→ `users.id`), `created_at`.

При первом запуске создаётся 5 тестовых заметок.

## Ответ `GET /api/notes`

```json
[
  {
    "id": "00001",
    "title_user": "Конференция ИТ - user25",
    "content": "Содержимое заметки",
    "formatted_date": "15.04.2027"
  }
]
```

- `id` — код заметки;
- `title_user` — `"{title} - {login}"`;
- `content` — содержание без изменений;
- `formatted_date` — дата в формате `ДД.ММ.ГГГГ`.

## Ошибки

| Ситуация | Статус | Ответ |
|----------|--------|-------|
| Плохой параметр `id_user` | 400 | `{"error": "..."}` |
| Записей нет | 200 | `[]` |
| Сбой БД | 500 | `{"error": "Ошибка подключения к базе данных"}` |

## Запуск

```sh
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --reload
```

Проверка: <http://127.0.0.1:8000/api/notes>

## Тесты

```sh
pytest -v
```

Postman: импортируйте `postman_collection.json` (переменная `base_url` =
`http://127.0.0.1:8000`). Тест «500» использует `?simulate_failure=true`.

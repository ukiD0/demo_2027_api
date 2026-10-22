import sqlite3
from datetime import datetime

from fastapi import *
from fastapi.responses import JSONResponse
from sqlite_utils import Database

app = FastAPI()
db = Database(sqlite3.connect("notes.db", check_same_thread=False))


def format_date(value) -> str:
    """Преобразует дату из базы в строку формата ДД.ММ.ГГГГ.

    Прил_5, п. 2.2: поле formatted_date = created_at в формате ДД.ММ.ГГГГ.
    Пример: 2027-04-15  ->  15.04.2027
    """
    if value is None:
        return ""

    text = str(value).strip()
    if not text:
        return ""

    # Форматы, в которых дата может лежать в базе.
    patterns = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d",
        "%d.%m.%Y",
    )
    for pattern in patterns:
        try:
            return datetime.strptime(text, pattern).strftime("%d.%m.%Y")
        except ValueError:
            continue

    # Запасной вариант: ISO-строка с миллисекундами или часовым поясом.
    # Берём первые 10 символов вида ГГГГ-ММ-ДД.
    if len(text) >= 10 and text[4] == "-" and text[7] == "-":
        try:
            return datetime.strptime(text[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
        except ValueError:
            pass

    # Формат неизвестен — отдаём как есть, чтобы не потерять данные.
    return text


@app.get("/api/notes")
def get_notes(id_user: str | None = Query(None)):
    """Список заметок. 200 — ок, 400 — плохой параметр, 500 — сбой сервера."""
    if id_user and not id_user.isdigit():
        return JSONResponse(
            status_code=400,
            content={"error": "Параметр id_user должен быть целым числом"},
        )

    try:
        sql = """
            SELECT * FROM notes
            JOIN users ON users.id = notes.id_user
        """
        if id_user:
            sql += f"WHERE notes.id_user = {int(id_user)}"

        rows = db.query(sql)

        resp_arr = []
        for row in rows:
            resp_arr.append(
                {
                    "id": f"{row['id']:05d}",
                    "title_user": f"{row['title']} - {row['login']}",
                    "content": row["content"],
                    "formatted_date": format_date(row["created_at"]),
                }
            )
        return resp_arr

    except Exception as error:
        print("ОШИБКА при работе с базой данных:", error)
        return JSONResponse(
            status_code=500,
            content={"error": "Ошибка подключения к базе данных"},
        )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
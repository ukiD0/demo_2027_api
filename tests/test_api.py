"""Тесты FastAPI-приложения (проверяют 200, 400, 500 и валидность JSON)."""
import os
import sys

import pytest
from fastapi.testclient import TestClient

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module


@pytest.fixture()
def client():
    """Тестовый клиент приложения."""
    return TestClient(app_module.app)


def test_status_200(client):
    """Штатная работа: 200 OK."""
    assert client.get("/api/notes").status_code == 200


def test_json_valid(client):
    """Ответ — валидный JSON-массив из 5 заметок."""
    data = client.get("/api/notes").json()
    assert isinstance(data, list) and len(data) == 5


def test_response_fields(client):
    """Поля и формат ответа совпадают с заданием."""
    assert client.get("/api/notes").json()[0] == {
        "id": "00001",
        "title_user": "Конференция ИТ - user25",
        "content": "Содержимое заметки",
        "formatted_date": "15.04.2027",
    }


def test_status_400(client):
    """Некорректный параметр -> 400 Bad Request."""
    response = client.get("/api/notes?id_user=abc")
    assert response.status_code == 400 and "error" in response.json()


def test_empty_result(client):
    """Нет записей -> 200 OK и пустой массив."""
    response = client.get("/api/notes?id_user=999")
    assert response.status_code == 200 and response.json() == []


def test_status_500(client):
    """Внутренняя ошибка -> 500 Internal Server Error."""
    response = client.get("/api/notes?simulate_failure=true")
    assert response.status_code == 500 and "error" in response.json()

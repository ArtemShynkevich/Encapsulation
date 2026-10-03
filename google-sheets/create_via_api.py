#!/usr/bin/env python3
"""Создаёт таблицу через Google Sheets API (нужен GOOGLE_SERVICE_ACCOUNT_JSON)."""
import json
import os
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

TITLE = "Instagram-лиды fakechat.pro"


def main() -> int:
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        print(
            "Задайте переменную GOOGLE_SERVICE_ACCOUNT_JSON (JSON сервисного аккаунта).",
            file=sys.stderr,
        )
        return 1

    info = json.loads(raw)
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    sheets = build("sheets", "v4", credentials=creds)
    drive = build("drive", "v3", credentials=creds)

    body = {
        "properties": {"title": TITLE, "locale": "ru_RU", "timeZone": "Europe/Minsk"},
        "sheets": [
            {"properties": {"title": "Лиды", "index": 0}},
            {"properties": {"title": "Статистика", "index": 1}},
            {"properties": {"title": "Графики", "index": 2}},
            {"properties": {"title": "Шаблоны", "index": 3}},
            {"properties": {"title": "Настройки", "index": 4}},
        ],
    }
    ss = sheets.spreadsheets().create(body=body).execute()
    sid = ss["spreadsheetId"]
    url = f"https://docs.google.com/spreadsheets/d/{sid}/edit"

    email = info.get("client_email")
    if email:
        drive.permissions().create(
            fileId=sid,
            body={"type": "user", "role": "writer", "emailAddress": os.environ.get("GOOGLE_SHARE_EMAIL", email)},
            fields="id",
        ).execute()

    # Полная разметка — через Apps Script (надёжнее для диаграмм и валидации).
    # После создания пустой структуры запустите createInstagramLeadsSpreadsheet вручную
    # или скопируйте готовую таблицу из шаблона.
    print(url)
    print(
        "\nПримечание: для валидаций, условного форматирования и графиков "
        "используйте CreateInstagramLeadsSheet.gs в Apps Script (см. README.md).",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

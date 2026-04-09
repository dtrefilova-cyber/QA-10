import re

import gspread
import streamlit as st
from google.oauth2.service_account import Credentials


def connect_google():
    """Підключення до Google Sheets."""
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scope,
    )
    return gspread.authorize(creds)


def extract_sheet_id(sheet_value):
    """Повертає sheet id з повного URL або сирого значення."""
    if not sheet_value:
        return ""

    value = str(sheet_value).strip()
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", value)
    if match:
        return match.group(1)

    return value


def normalize_header(value):
    """Нормалізує назви колонок."""
    text = str(value or "")
    text = text.replace("\ufeff", "").replace("\u00a0", " ").strip().upper()
    text = re.sub(r"\s+", " ", text)
    return text


def load_managers_config(google_client, log_sheet_id, worksheet_name="MANAGERS"):
    """Зчитує список менеджерів і проєктів з технічного аркуша."""
    worksheet = google_client.open_by_key(log_sheet_id).worksheet(worksheet_name)
    values = worksheet.get_all_values()

    if not values:
        return {
            "managers": [],
            "headers": [],
            "header_row_index": None,
            "raw_rows_count": 0,
            "valid_rows_count": 0,
        }

    required_headers = {"MANAGERS_NAME", "PROJECT", "SHEET_ID"}
    header_row_index = None
    headers = []

    for idx, row in enumerate(values[:10]):
        normalized_row = [normalize_header(cell) for cell in row]
        if required_headers.issubset(set(normalized_row)):
            header_row_index = idx
            headers = normalized_row
            break

    if header_row_index is None:
        return {
            "managers": [],
            "headers": [normalize_header(cell) for cell in values[0]],
            "header_row_index": None,
            "raw_rows_count": max(len(values) - 1, 0),
            "valid_rows_count": 0,
        }

    rows = values[header_row_index + 1 :]

    def get_value(row, column_name):
        try:
            index = headers.index(column_name)
        except ValueError:
            return ""

        if index >= len(row):
            return ""

        return row[index]

    managers = []
    for row in rows:
        manager_name = str(get_value(row, "MANAGERS_NAME")).strip()
        project_name = str(get_value(row, "PROJECT")).strip()
        sheet_id = extract_sheet_id(get_value(row, "SHEET_ID"))

        if not manager_name or not project_name or not sheet_id:
            continue

        managers.append(
            {
                "manager_name": manager_name,
                "project": project_name,
                "sheet_id": sheet_id,
            }
        )

    return {
        "managers": managers,
        "headers": headers,
        "header_row_index": header_row_index,
        "raw_rows_count": len(rows),
        "valid_rows_count": len(managers),
    }


CRITERIA_ROWS = {
    "Встановлення контакту": 5,
    "Спроба презентації": 6,
    "Домовленість про наступний контакт": 7,
    "Пропозиція бонусу": 8,
    "Завершення розмови": 9,
    "Передзвон клієнту": 10,
    "Не додумувати": 11,
    "Якість мовлення": 12,
    "Професіоналізм": 13,
    "Оформлення картки": 14,
    "Робота із запереченнями": 15,
    "Утримання клієнта": 16,
}


def format_score_sheet(x):
    """Форматує оцінку для Google Sheets."""
    try:
        return float(x)
    except (ValueError, TypeError):
        return 0.0


def find_next_column(sheet):
    """Знаходить наступну вільну колонку для блоку оцінок."""
    try:
        row = sheet.row_values(3)
        for i, value in enumerate(row, start=1):
            if not value or value.strip() == "":
                return i
        return len(row) + 1
    except Exception:
        return 1


def find_next_row(sheet, start_row=1, key_column=1):
    """Знаходить перший вільний рядок, починаючи зі start_row."""
    try:
        column_values = sheet.col_values(key_column)
        row = start_row

        while row <= len(column_values):
            value = column_values[row - 1] if row - 1 < len(column_values) else ""
            if not str(value).strip():
                return row
            row += 1

        return max(start_row, len(column_values) + 1)
    except Exception:
        return start_row


def write_to_google_sheet(sheet, meta, scores):
    """Записує блок оцінок у таблицю менеджера по колонках."""
    try:
        column = find_next_column(sheet)

        def get_column_letter(n):
            string = ""
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                string = chr(65 + remainder) + string
            return string

        col_letter = get_column_letter(column)
        updates = [
            (f"{col_letter}1", meta.get("call_date", "")),
            (f"{col_letter}2", meta.get("client_id", "")),
            (f"{col_letter}3", meta.get("qa_manager", "")),
            (f"{col_letter}4", meta.get("check_date", "")),
        ]

        for key, value in scores.items():
            if key in CRITERIA_ROWS:
                row = CRITERIA_ROWS[key]
                updates.append((f"{col_letter}{row}", format_score_sheet(value)))

        for cell, val in updates:
            sheet.update(cell, [[val]])

        return True
    except Exception as e:
        return str(e)


def append_manager_log(sheet, call, comment, total_score, ai_label, start_row=20):
    """Додає підсумок перевірки в таблицю менеджера з рядка 20."""
    row_index = find_next_row(sheet, start_row=start_row, key_column=1)
    values = [[
        call.get("client_id", ""),
        comment,
        total_score,
        call.get("call_date", ""),
        call.get("check_date", ""),
        ai_label,
    ]]
    sheet.update(f"A{row_index}:F{row_index}", values, value_input_option="RAW")
    return row_index


def append_qa_log(sheet, call, transcript, clean_dialogue, comment, total_score):
    """Додає лог перевірки у QA_LOG_CALLS / Лист1."""
    row_index = find_next_row(sheet, start_row=1, key_column=1)
    values = [[
        call.get("check_date", ""),
        call.get("client_id", ""),
        call.get("project", ""),
        call.get("qa_manager", ""),
        call.get("url", ""),
        transcript,
        clean_dialogue,
        comment,
        total_score,
    ]]
    sheet.update(f"A{row_index}:I{row_index}", values, value_input_option="RAW")
    return row_index

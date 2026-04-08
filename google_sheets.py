import streamlit as st
import gspread
import re
from google.oauth2.service_account import Credentials


def connect_google():
    """Підключення до Google Sheets"""
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scope
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
    """Нормалізує назви колонок: прибирає BOM, зайві пробіли та службові символи."""
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
            "valid_rows_count": 0
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
            "valid_rows_count": 0
        }

    rows = values[header_row_index + 1:]

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

        managers.append({
            "manager_name": manager_name,
            "project": project_name,
            "sheet_id": sheet_id
        })

    return {
        "managers": managers,
        "headers": headers,
        "header_row_index": header_row_index,
        "raw_rows_count": len(rows),
        "valid_rows_count": len(managers)
    }


# 🔹 Маппінг критеріїв (де і як вносяться оцінки)
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
    "Утримання клієнта": 16
}

META_ROWS = {
    "call_date": 1,
    "qa_manager": 2,
    "client_id": 3,
    "check_date": 4
}


def format_score_sheet(x):
    """Форматує оцінку для Google Sheets"""
    try:
        return float(x)
    except (ValueError, TypeError):
        return 0.0


def find_next_column(sheet):
    """Знаходить наступну вільну колонку"""
    try:
        row = sheet.row_values(3)  # рядок client_id
        for i, value in enumerate(row, start=1):
            if not value or value.strip() == "":
                return i
        return len(row) + 1
    except:
        return 1


def write_to_google_sheet(sheet, meta, scores):
    """Записує результати в Google Sheets"""

    try:
        column = find_next_column(sheet)
        def get_column_letter(n):
            string = ""
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                string = chr(65 + remainder) + string
            return string
        
        col_letter = get_column_letter(column)

        updates = []

        # 🔹 мета-дані (рядки 1–4)
        updates.extend([
            (f"{col_letter}1", meta.get("call_date", "")),     # 1 — дата дзвінка
            (f"{col_letter}2", meta.get("client_id", "")),     # 2 — айді клієнта
            (f"{col_letter}3", meta.get("qa_manager", "")),    # 3 — QA
            (f"{col_letter}4", meta.get("check_date", ""))     # 4 — дата перевірки
        ])

        # 🔹 оцінки (рядки 5+)
        for key, value in scores.items():
            if key in CRITERIA_ROWS:
                row = CRITERIA_ROWS[key]
                updates.append((f"{col_letter}{row}", format_score_sheet(value)))

        # 🔹 запис у таблицю
        for cell, val in updates:
            sheet.update(cell, [[val]])

        return True

    except Exception as e:
        return str(e)
   
    # Метадані
    updates.append((META_ROWS["call_date"], meta.get("call_date", "")))
    updates.append((META_ROWS["qa_manager"], meta.get("qa_manager", "")))
    updates.append((META_ROWS["client_id"], meta.get("client_id", "")))
    updates.append((META_ROWS["check_date"], meta.get("check_date", "")))
   
    # Оцінки
    for criterion, score in scores.items():
        if criterion in CRITERIA_ROWS:
            row = CRITERIA_ROWS[criterion]
            updates.append((row, format_score_sheet(score)))
        else:
            # Якщо з'явився новий критерій, який ще не в маппінгу — можна логувати
            print(f"Попередження: Критерій '{criterion}' не знайдено в CRITERIA_ROWS")
   
    # Записуємо в таблицю
    cell_list = [gspread.Cell(row, column, value) for row, value in updates]
    sheet.update_cells(cell_list)

import streamlit as st
import gspread
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
        col_letter = chr(64 + column)

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

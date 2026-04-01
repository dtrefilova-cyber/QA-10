# google_sheets.py
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

# Маппінг критеріїв на рядки в таблиці
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
    """Форматує оцінку для Google Sheets - повертає число"""
    return float(x)

def find_next_column(sheet):
    """Знаходить наступну вільну колонку"""
    row = sheet.row_values(META_ROWS["client_id"])
    for i, value in enumerate(row, start=1):
        if value == "":
            return i
    return len(row) + 1

def write_to_google_sheet(sheet, meta, scores):
    """Записує результати в Google Sheets"""
    column = find_next_column(sheet)
    updates = []
    
    # Метадані
    updates.append((META_ROWS["call_date"], meta["call_date"]))
    updates.append((META_ROWS["qa_manager"], meta["qa_manager"]))
    updates.append((META_ROWS["client_id"], meta["client_id"]))
    updates.append((META_ROWS["check_date"], meta["check_date"]))
    
    # Оцінки
    for criterion, score in scores.items():
        if criterion in CRITERIA_ROWS:
            row = CRITERIA_ROWS[criterion]
            updates.append((row, format_score_sheet(score)))
    
    # Записуємо в таблицю
    cell_list = []
    for row, value in updates:
        cell_list.append(gspread.Cell(row, column, value))
    sheet.update_cells(cell_list)

import streamlit as st
import pandas as pd
import requests
import json
import re
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
from datetime import datetime
from openai import OpenAI

DEEPGRAM_API_KEY = st.secrets["DEEPGRAM_API_KEY"]
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]

client = OpenAI(api_key=OPENAI_API_KEY)

st.title("🎧 QA-10: Аналіз дзвінків")

check_date = st.date_input("Дата перевірки", datetime.today())

qa_managers_list = [
    "Аліна",
    "Дар'я",
    "Надя",
    "Настя",
    "Владимира",
    "Діана",
    "Руслана",
    "Олексій"
]

calls = []

for row in range(5):
    col1, col2 = st.columns(2)
    for col, idx in zip([col1, col2], [row * 2 + 1, row * 2 + 2]):
        with col.expander(f"📞 Дзвінок {idx}", expanded=False):
            audio_url = st.text_input("Посилання на аудіо", key=f"url_{idx}")
            qa_manager = st.selectbox("QA менеджер", qa_managers_list, key=f"qa_{idx}")
            ret_manager = st.text_input("Менеджер RET", key=f"ret_{idx}")
            client_id = st.text_input("ID клієнта", key=f"client_{idx}")
            call_date = st.text_input("Дата дзвінка (ДД-ММ-РРРР)", key=f"date_{idx}")
            bonus_check = st.selectbox("Бонус",
                ["правильно нараховано", "помилково нараховано", "не потрібно"],
                key=f"bonus_{idx}")
            repeat_call = st.selectbox("Повторний дзвінок",
                ["так, був протягом години", "так, був протягом 3 годин", "ні, не було"],
                key=f"repeat_{idx}")
            manager_comment = st.text_area("Коментар менеджера", height=80, key=f"comment_{idx}")
            speech_score = st.selectbox("Якість мовлення (ручна оцінка)", [2.5, 0], key=f"speech_{idx}")

            calls.append({
                "url": audio_url,
                "qa_manager": qa_manager,
                "ret_manager": ret_manager,
                "client_id": client_id,
                "call_date": call_date,
                "check_date": check_date.strftime("%d-%m-%Y"),
                "bonus_check": bonus_check,
                "repeat_call": repeat_call,
                "manager_comment": manager_comment,
                "speech_score": speech_score
            })

# ---------------- GOOGLE SHEETS ----------------

def connect_google():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scope
    )
    return gspread.authorize(creds)

CRITERIA_ROWS = {
    "Привітання": 5,
    "Дружелюбне питання / Мета дзвінка": 6,
    "Спроба продовжити розмову": 7,
    "Спроба презентації": 8,
    "Домовленість про наступний контакт": 9,
    "Пропозиція бонусу": 10,
    "Завершення": 11,
    "Передзвон клієнту": 12,
    "Не додумувати": 13,
    "Якість мовлення": 14,
    "Професіоналізм": 15,
    "CRM-картка": 16,
    "Робота із запереченнями": 17,
    "Зливання клієнта": 18
}

META_ROWS = {
    "call_date": 1,
    "qa_manager": 2,
    "client_id": 3,
    "check_date": 4
}

def find_next_column(sheet):
    row = sheet.row_values(META_ROWS["client_id"])
    for i, value in enumerate(row, start=1):
        if value == "":
            return i
    return len(row) + 1

def format_score(x):
    return f"{float(x):.1f}"

def format_score_sheet(x):
    return format_score(x).replace(".", ",")

def write_to_google_sheet(sheet, meta, scores):
    column = find_next_column(sheet)
    updates = []
    updates.append((META_ROWS["call_date"], meta["call_date"]))
    updates.append((META_ROWS["qa_manager"], meta["qa_manager"]))
    updates.append((META_ROWS["client_id"], meta["client_id"]))
    updates.append((META_ROWS["check_date"], meta["check_date"]))
    for criterion, score in scores.items():
        if criterion in CRITERIA_ROWS:
            row = CRITERIA_ROWS[criterion]
            updates.append((row, format_score_sheet(score)))
    cell_list = [gspread.Cell(row, column, value) for row, value in updates]
    sheet.update_cells(cell_list)

# ---------------- TRANSCRIPTION, FEATURES, COMMENT, SCORING ----------------
# (залишаємо весь блок із файлу 4 без змін)

# ---------------- ANALYSIS ----------------

if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("Запустити аналіз"):
    st.session_state["results"].clear()
    google_client = connect_google()
    for i, call in enumerate(calls):
        if not call["url"]:
            continue
        st.write(f"⏳ Обробка дзвінка {i+1}...")
        transcript = transcribe_audio(call["url"])
        if not transcript:
            st.write("⚠️ Не вдалося отримати транскрипт")
            continue
        features = extract_features(transcript)
        scores = score_call(features, call)
        comment = generate_comment(transcript)
        try:
            spreadsheet = google_client.open(call["ret_manager"])
            sheet = spreadsheet.sheet1
            write_to_google_sheet(sheet, call, scores)
        except Exception as e:
            st.warning(f"Google Sheets error: {e}")
        st.session_state["results"].append({
            "meta": call,
            "scores": scores,
            "comment": comment
        })

# ---------------- OUTPUT ----------------

for i, res in enumerate(st.session_state["results"]):
    with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):
        df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(format_score)

        st.table(df)

        total_score = f"{sum(res['scores'].values()):.1f}"
        st.markdown(f"**Загальний бал:** {total_score}")

        st.markdown("### Коментар")
        st.write(res["comment"])

# ---------------- EXPORT EXCEL ----------------

if st.session_state["results"]:
    xls = BytesIO()
    with pd.ExcelWriter(xls, engine="openpyxl") as writer:
        for i, res in enumerate(st.session_state["results"]):
            sheet_name = f"Call_{i+1}"

            # Метадані
            meta_df = pd.DataFrame(list(res["meta"].items()), columns=["Поле", "Значення"])
            meta_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)

            # Оцінки
            scores_df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
            scores_df["Оцінка"] = scores_df["Оцінка"].apply(format_score)
            scores_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=len(meta_df) + 2)

            # Коментар
            comment_df = pd.DataFrame([["Коментар", res["comment"]]], columns=["Поле", "Значення"])
            comment_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=len(meta_df) + len(scores_df) + 4)

    xls.seek(0)

    st.download_button(
        "📥 Завантажити результати у XLSX",
        xls,
        "qa_results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


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
    "Аліна Пронь",
    "Дар'я Трефілова",
    "Надія Татаренко",
    "Анастасія Собакіна",
    "Владимира Балховська",
    "Діана Батрак",
    "Руслана Каленіченко",
    "Шутов Олексій"
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

# ---------------- TRANSCRIPTION ----------------

def transcribe_audio(audio_url):
    if not audio_url:
        return None
    url = "https://api.deepgram.com/v1/listen"
    params = {
        "model": "nova-2",
        "language": "uk",
        "diarize": "true",
        "utterances": "true",
        "punctuate": "true",
        "smart_format": "true"
    }
    headers = {"Authorization": f"Token {DEEPGRAM_API_KEY}"}
    response = requests.post(url, headers=headers, params=params, json={"url": audio_url})
    if response.status_code != 200:
        return None
    result = response.json()
    if "results" not in result:
        return None
    dialogue = []
    current_speaker = None
    current_text = ""
    for u in result["results"]["utterances"]:
        speaker = "Менеджер" if u["speaker"] == 0 else "Гравець"
        text = u["transcript"].strip()
        if speaker == current_speaker:
            current_text += " " + text
        else:
            if current_speaker is not None:
                dialogue.append(f"{current_speaker}: {current_text}")
            current_speaker = speaker
            current_text = text
    if current_text:
        dialogue.append(f"{current_speaker}: {current_text}")
    return "\n".join(dialogue)

# ---------------- SEGMENTS ----------------

def extract_segments(dialogue):
    lines = dialogue.split("\n")
    intro = "\n".join(lines[:4])
    middle = "\n".join(lines[4:-4])
    ending = "\n".join(lines[-4:])
    return intro, middle, ending

# ---------------- FEATURE EXTRACTION ----------------

def extract_features(dialogue):
    intro, middle, ending = extract_segments(dialogue)
    prompt = f"""
Проаналізуй дзвінок...
(тут правила як у файлі 4)
Початок:
{intro}
Середина:
{middle}
Кінець:
{ending}
Поверни тільки JSON.
{{ ... }}
"""
    response = client.chat.completions.create(
        model="gpt-4.1",
        temperature=0,
        messages=[
            {"role": "system", "content": "Ти система аналізу дзвінків."},
            {"role": "user", "content": prompt}
        ]
    )
    text = response.choices[0].message.content
    match = re.search(r"\{.*\}", text, re.S)
    if match:
        features = json.loads(match.group())
    else:
        features = {}
    defaults = {
        "manager_introduced_self": False,
        "client_name_used": False,
        "presentation_detected": False,
        "bonus_offered": False,
        "bonus_conditions_count": 0,
        "client_busy": False,
        "manager_active": True,
        "followup_type": "none",
        "objection_detected": False
    }
    for k, v in defaults.items():
        features.setdefault(k, v)
    features["raw_text"] = dialogue.lower()
    return features

# ---------------- COMMENT ----------------

def generate_comment(dialogue):
    prompt = f"""
Коротко підсумуй дзвінок менеджера казино.
1–2 речення. Вкажи сильну сторону і рекомендацію.
{dialogue}
"""
    response = client.chat.completions.create(
        model="gpt-4.1",
        temperature=0.3,
        messages=[
            {"role": "system", "content": "Ти QA-аналітик дзвінків."},
            {"role": "user", "content": prompt}
        ]
    )
    return response.choices[0].message.content

# ---------------- SCORING ----------------

def score_call(features, meta):
    scores = {}

    introduced = features["manager_introduced_self"]
    client_name = features["client_name_used"]

    if introduced and client_name:
        scores["Привітання"] = 5
    elif introduced or client_name:
        scores["Привітання"] = 2.5
    else:
        scores["Привітання"] = 0

    scores["Дружелюбне питання / Мета дзвінка"] = 2.5
    scores["Спроба продовжити розмову"] = 5 if features["manager_active"] else 0

    presentation = features["presentation_detected"]
    if not presentation:
        text = features.get("raw_text", "")
        if any(word in text for word in ["слот", "бонус", "активн", "турнір", "фріспін"]):
            presentation = True
    scores["Спроба презентації"] = 5 if presentation else 0

    followup = features["followup_type"]
    if followup == "exact_time":
        scores["Домовленість про наступний контакт"] = 10
    elif followup == "day":
        scores["Домовленість про наступний контакт"] = 7.5
    elif followup == "offer":
        scores["Домовленість про наступний контакт"] = 5
    else:
        scores["Домовленість про наступний контакт"] = 0

    bonus_conditions = features["bonus_conditions_count"]
    if not features["bonus_offered"]:
        scores["Пропозиція бонусу"] = 0
    elif bonus_conditions == 0:
        scores["Пропозиція бонусу"] = 5
    elif bonus_conditions == 1:
        scores["Пропозиція бонусу"] = 7.5
    else:
        scores["Пропозиція бонусу"] = 10

    if features["client_busy"]:
        scores["Завершення"] = 5
    elif features["manager_active"]:
        scores["Завершення"] = 5
    else:
        scores["Завершення"] = 0

    repeat = meta["repeat_call"]
    if repeat == "так, був протягом години":
        scores["Передзвон клієнту"] = 10
    elif repeat == "так, був протягом 3 годин":
        scores["Передзвон клієнту"] = 5
    else:
        if followup != "none":
            scores["Передзвон клієнту"] = 0
        else:
            scores["Передзвон клієнту"] = 10

    scores["Не додумувати"] = 5
    scores["Якість мовлення"] = meta["speech_score"]
    scores["Професіоналізм"] = 5 if meta["bonus_check"] == "помилково нараховано" else 10
    scores["CRM-картка"] = 5 if meta["manager_comment"] else 0
    scores["Робота із запереченнями"] = 10 if not features["objection_detected"] else 5

    if features["client_busy"]:
        scores["Зливання клієнта"] = 15
    elif features["manager_active"]:
        scores["Зливання клієнта"] = 15
    else:
        scores["Зливання клієнта"] = 10

    return scores

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

            meta_df = pd.DataFrame(list(res["meta"].items()), columns=["Поле", "Значення"])
            meta_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)

            scores_df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
            scores_df["Оцінка"] = scores_df["Оцінка"].apply(format_score)
            scores_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=len(meta_df) + 2)

            comment_df = pd.DataFrame([["Коментар", res["comment"]]], columns=["Поле", "Значення"])
            comment_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=len(meta_df) + len(scores_df) + 4)

    xls.seek(0)
    st.download_button(
        "📥 Завантажити результати у XLSX",
        xls,
        "qa_results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

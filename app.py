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


def write_to_google_sheet(meta, scores):

    try:

        client_g = connect_google()

        sheet = client_g.open("QA_RESULTS").sheet1

        row = [
            meta["check_date"],
            meta["qa_manager"],
            meta["ret_manager"],
            meta["client_id"],
            meta["call_date"],
        ]

        for k in scores:
            row.append(float(scores[k]))

        sheet.append_row(row)

    except Exception as e:

        st.warning(f"Помилка запису в Google Sheets: {e}")


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

            qa_manager = st.selectbox(
                "QA менеджер",
                qa_managers_list,
                key=f"qa_{idx}"
            )

            ret_manager = st.text_input("Менеджер RET", key=f"ret_{idx}")

            client_id = st.text_input("ID клієнта", key=f"client_{idx}")

            call_date = st.text_input(
                "Дата дзвінка (ДД-ММ-РРРР)",
                key=f"date_{idx}"
            )

            bonus_check = st.selectbox(
                "Бонус",
                ["правильно нараховано", "помилково нараховано", "не потрібно"],
                key=f"bonus_{idx}"
            )

            repeat_call = st.selectbox(
                "Повторний дзвінок",
                [
                    "так, був протягом години",
                    "так, був протягом 3 годин",
                    "ні, не було"
                ],
                key=f"repeat_{idx}"
            )

            manager_comment = st.text_area(
                "Коментар менеджера",
                height=80,
                key=f"comment_{idx}"
            )

            speech_score = st.selectbox(
                "Якість мовлення (ручна оцінка)",
                [2.5, 0],
                key=f"speech_{idx}"
            )

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

    response = requests.post(
        url,
        headers=headers,
        params=params,
        json={"url": audio_url}
    )

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
Проаналізуй дзвінок.

Поверни тільки JSON.

{{
"manager_introduced_self": true/false,
"client_name_used": true/false,
"presentation_detected": true/false,
"bonus_offered": true/false,
"bonus_conditions_count": number,
"client_busy": true/false,
"manager_active": true/false,
"followup_type": "none / offer / day / exact_time",
"objection_detected": true/false
}}

{dialogue}
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
        return json.loads(match.group())

    return {}


# ---------------- COMMENT ----------------

def generate_comment(dialogue):

    prompt = f"""
Коротко підсумуй дзвінок менеджера.

1–2 речення.

{dialogue}
"""

    response = client.chat.completions.create(
        model="gpt-4.1",
        temperature=0.3,
        messages=[
            {"role": "system", "content": "Ти QA-аналітик."},
            {"role": "user", "content": prompt}
        ]
    )

    return response.choices[0].message.content


# ---------------- SCORING ----------------

def score_call(features, meta):

    scores = {}

    introduced = features.get("manager_introduced_self", False)
    client_name = features.get("client_name_used", False)

    if introduced and client_name:
        scores["Привітання"] = 5
    elif introduced or client_name:
        scores["Привітання"] = 2.5
    else:
        scores["Привітання"] = 0

    scores["Дружелюбне питання / Мета дзвінка"] = 2.5

    scores["Спроба продовжити розмову"] = 5 if features.get("manager_active") else 0

    scores["Спроба презентації"] = 5 if features.get("presentation_detected") else 0

    scores["Домовленість про наступний контакт"] = 5

    scores["Пропозиція бонусу"] = 5 if features.get("bonus_offered") else 0

    scores["Завершення"] = 2.5

    scores["Передзвон клієнту"] = 10

    scores["Не додумувати"] = 5

    scores["Якість мовлення"] = meta["speech_score"]

    scores["Професіоналізм"] = 10

    scores["CRM-картка"] = 5 if meta["manager_comment"] else 0

    scores["Робота із запереченнями"] = 10

    scores["Зливання клієнта"] = 15

    return scores


# ---------------- ANALYSIS ----------------

if st.button("Запустити аналіз"):

    for call in calls:

        if not call["url"]:
            continue

        transcript = transcribe_audio(call["url"])

        if not transcript:
            continue

        features = extract_features(transcript)

        scores = score_call(features, call)

        comment = generate_comment(transcript)

        write_to_google_sheet(call, scores)

        st.write(scores)
        st.write(comment)

        # ---------- ВИВІД ТАБЛИЦІ ----------

        df = pd.DataFrame(scores.items(), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].astype(float)

        st.table(df)

        total_score = round(sum(scores.values()), 1)

        st.markdown(f"**Загальний бал:** {total_score}")

        st.markdown("### Коментар")
        st.write(comment)

        # ---------- ЗБЕРЕЖЕННЯ В SESSION ----------

        if "results" not in st.session_state:
            st.session_state["results"] = []

        st.session_state["results"].append({
            "meta": call,
            "scores": scores,
            "comment": comment
        })


# ---------------- EXPORT XLSX ----------------

if "results" in st.session_state and st.session_state["results"]:

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

        for i, res in enumerate(st.session_state["results"]):

            sheet_name = f"Call_{i+1}"

            meta_df = pd.DataFrame(
                list(res["meta"].items()),
                columns=["Поле", "Значення"]
            )

            scores_df = pd.DataFrame(
                list(res["scores"].items()),
                columns=["Критерій", "Оцінка"]
            )

            scores_df["Оцінка"] = scores_df["Оцінка"].astype(float)

            meta_df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=0,
                index=False
            )

            scores_df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=len(meta_df) + 2,
                index=False
            )

            comment_df = pd.DataFrame(
                [["Коментар", res["comment"]]],
                columns=["Поле", "Значення"]
            )

            comment_df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=len(meta_df) + len(scores_df) + 4,
                index=False
            )

    buffer.seek(0)

    st.download_button(
        "📥 Завантажити результати (XLSX)",
        buffer,
        "qa_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

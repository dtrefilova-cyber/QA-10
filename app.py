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


# ---------------- GOOGLE ----------------

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


META_ROWS = {
    "call_date": 1,
    "qa_manager": 2,
    "client_id": 3,
    "check_date": 4
}


CRITERIA_ROWS = {

"Привітання":5,
"Дружелюбне питання / Мета дзвінка":6,
"Спроба продовжити розмову":7,
"Спроба презентації":8,
"Домовленість про наступний контакт":9,
"Пропозиція бонусу":10,
"Завершення":11,

"Передзвон клієнту":12,
"Не додумувати":13,
"Якість мовлення":14,
"Професіоналізм":15,
"CRM-картка":16,
"Робота із запереченнями":17,
"Зливання клієнта":18
}


def find_next_column(sheet):

    row = sheet.row_values(META_ROWS["client_id"])

    for i, value in enumerate(row, start=1):

        if value == "":
            return i

    return len(row) + 1


def find_manager_sheet(client, manager_name):

    manager_name = manager_name.lower().strip()

    for file in client.openall():

        if manager_name in file.title.lower():
            return file

    return None


def write_to_google_sheet(client, meta, scores):

    spreadsheet = find_manager_sheet(client, meta["ret_manager"])

    if not spreadsheet:
        st.warning(f"Таблицю менеджера '{meta['ret_manager']}' не знайдено")
        return

    sheet = spreadsheet.sheet1
    column = find_next_column(sheet)

    updates = []

    updates.append((META_ROWS["call_date"], meta["call_date"]))
    updates.append((META_ROWS["qa_manager"], meta["qa_manager"]))
    updates.append((META_ROWS["client_id"], meta["client_id"]))
    updates.append((META_ROWS["check_date"], meta["check_date"]))

    for criterion, score in scores.items():

        if criterion in CRITERIA_ROWS:

            row = CRITERIA_ROWS[criterion]
            updates.append((row, float(score)))

    cells = []

    for row, value in updates:
        cells.append(gspread.Cell(row, column, value))

    sheet.update_cells(cells)


# ---------------- TRANSCRIPTION ----------------

def transcribe_audio(audio_url):

    if not audio_url:
        return None

    url = "https://api.deepgram.com/v1/listen"

    params = {
        "model": "nova-2",
        "language": "uk",
        "diarize": True,
        "utterances": True,
        "punctuate": True
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

    try:

        dialogue = []
        current_speaker = None
        current_text = ""

        for u in result["results"]["utterances"]:

            speaker = "Менеджер" if u["speaker"] == 0 else "Клієнт"
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

    except Exception:
        return None


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
Проаналізуй дзвінок менеджера.

Початок:
{intro}

Середина:
{middle}

Кінець:
{ending}

Поверни JSON:

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
"""

    response = client.chat.completions.create(
        model="gpt-4.1",
        temperature=0,
        messages=[
            {"role":"system","content":"Ти система аналізу дзвінків"},
            {"role":"user","content":prompt}
        ]
    )

    text = response.choices[0].message.content

    match = re.search(r"\{.*\}", text, re.S)

    if match:
        features = json.loads(match.group())
    else:
        features = {}

    return features


# ---------------- COMMENT ----------------

def generate_comment(dialogue):

    prompt = f"""
Коротко підсумуй дзвінок менеджера.

1-2 речення: сильна сторона + рекомендація.

{dialogue}
"""

    response = client.chat.completions.create(
        model="gpt-4.1",
        temperature=0.3,
        messages=[
            {"role":"system","content":"Ти QA-аналітик"},
            {"role":"user","content":prompt}
        ]
    )

    return response.choices[0].message.content


# ---------------- SCORING ----------------

def score_call(features, meta):

    scores = {}

    introduced = features.get("manager_introduced_self")
    client_name = features.get("client_name_used")

    if introduced and client_name:
        scores["Привітання"] = 5
    elif introduced or client_name:
        scores["Привітання"] = 2.5
    else:
        scores["Привітання"] = 0

    scores["Дружелюбне питання / Мета дзвінка"] = 2.5
    scores["Спроба продовжити розмову"] = 5 if features.get("manager_active") else 0
    scores["Спроба презентації"] = 5 if features.get("presentation_detected") else 0

    followup = features.get("followup_type")

    if followup == "exact_time":
        scores["Домовленість про наступний контакт"] = 10
    elif followup == "day":
        scores["Домовленість про наступний контакт"] = 7.5
    elif followup == "offer":
        scores["Домовленість про наступний контакт"] = 5
    else:
        scores["Домовленість про наступний контакт"] = 0

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

    scores["Робота із запереченнями"] = 10 if not features.get("objection_detected") else 5
    scores["Зливання клієнта"] = 15

    return scores


# ---------------- UI ----------------

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

    for col, idx in zip([col1, col2], [row*2+1, row*2+2]):

        with col.expander(f"📞 Дзвінок {idx}"):

            audio_url = st.text_input("Посилання на аудіо", key=f"url_{idx}")

            qa_manager = st.selectbox("QA менеджер", qa_managers_list, key=f"qa_{idx}")

            ret_manager = st.text_input("Менеджер RET", key=f"ret_{idx}")

            client_id = st.text_input("ID клієнта", key=f"client_{idx}")

            call_date = st.text_input("Дата дзвінка", key=f"date_{idx}")

            bonus_check = st.selectbox(
                "Бонус",
                ["правильно нараховано","помилково нараховано","не потрібно"],
                key=f"bonus_{idx}"
            )

            repeat_call = st.selectbox(
                "Повторний дзвінок",
                ["так, був протягом години","так, був протягом 3 годин","ні, не було"],
                key=f"repeat_{idx}"
            )

            manager_comment = st.text_area("Коментар менеджера",key=f"comment_{idx}")

            speech_score = st.selectbox(
                "Якість мовлення",
                [2.5,0],
                key=f"speech_{idx}"
            )

            calls.append({
                "url":audio_url,
                "qa_manager":qa_manager,
                "ret_manager":ret_manager,
                "client_id":client_id,
                "call_date":call_date,
                "check_date":check_date.strftime("%d-%m-%Y"),
                "bonus_check":bonus_check,
                "repeat_call":repeat_call,
                "manager_comment":manager_comment,
                "speech_score":speech_score
            })


status = st.empty()
progress = st.progress(0)


# ---------------- ANALYSIS ----------------

if st.button("Запустити аналіз"):

    results = []

    google_client = connect_google()

    total = len([c for c in calls if c["url"]])
    done = 0

    for i, call in enumerate(calls):

        if not call["url"]:
            continue

        status.write(f"Аналіз дзвінка {i+1}")

        transcript = transcribe_audio(call["url"])

        if not transcript:
            status.write("⚠️ Транскрипцію не отримано")
            continue

        features = extract_features(transcript)
        scores = score_call(features, call)
        comment = generate_comment(transcript)

        write_to_google_sheet(google_client, call, scores)

        results.append({
            "meta":call,
            "scores":scores,
            "comment":comment
        })

        done += 1
        progress.progress(done/total)

    st.session_state["results"] = results

    status.success("Аналіз завершено")


# ---------------- RESULTS ----------------

if "results" in st.session_state:

    for i,res in enumerate(st.session_state["results"]):

        with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):

            df = pd.DataFrame(res["scores"].items(),columns=["Критерій","Оцінка"])

            st.table(df)

            total = sum(res["scores"].values())

            st.markdown(f"### Загальний бал: {total}")

            st.markdown("### Коментар QA")
            st.write(res["comment"])


# ---------------- EXPORT ----------------

if "results" in st.session_state:

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

        for i,res in enumerate(st.session_state["results"]):

            name=f"Call_{i+1}"

            meta_df=pd.DataFrame(list(res["meta"].items()),columns=["Поле","Значення"])
            meta_df.to_excel(writer,index=False,sheet_name=name)

            scores_df=pd.DataFrame(res["scores"].items(),columns=["Критерій","Оцінка"])
            scores_df.to_excel(writer,index=False,sheet_name=name,startrow=len(meta_df)+2)

            comment_df=pd.DataFrame([["Коментар",res["comment"]]],columns=["Поле","Значення"])
            comment_df.to_excel(writer,index=False,sheet_name=name,startrow=len(meta_df)+len(scores_df)+4)

    buffer.seek(0)

    st.download_button(
        "📥 Завантажити результати XLSX",
        buffer,
        "qa_results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

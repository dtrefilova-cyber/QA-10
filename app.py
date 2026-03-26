import streamlit as st
import pandas as pd
import requests
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
from datetime import datetime
from openai import OpenAI


DEEPGRAM_API_KEY = st.secrets["DEEPGRAM_API_KEY"]
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]

client = OpenAI(api_key=OPENAI_API_KEY)


# ---------------- FORMAT ----------------

def format_score(x):
    return f"{float(x):.1f}"

def format_score_sheet(x):
    return format_score(x).replace(".", ",")


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


# ---------------- FIND COLUMN ----------------

def find_next_column(sheet):

    row = sheet.row_values(META_ROWS["client_id"])

    for i, value in enumerate(row, start=1):

        if value == "":
            return i

    return len(row) + 1


# ---------------- FIND TABLE ----------------

def find_manager_sheet(client, manager_name):

    manager_name = manager_name.lower().strip()

    for file in client.openall():

        if manager_name in file.title.lower():
            return file

    return None


# ---------------- WRITE GOOGLE ----------------

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

            updates.append((row, format_score_sheet(score)))

    cells = []

    for row, value in updates:

        cells.append(gspread.Cell(row, column, value))

    sheet.update_cells(cells)


# ---------------- TRANSCRIBE ----------------

def transcribe_audio(audio_url):

    url = "https://api.deepgram.com/v1/listen"

    params = {
        "model": "nova-2",
        "language": "uk",
        "diarize": True,
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

        if "utterances" in result["results"]:

            dialogue = []

            for u in result["results"]["utterances"]:
                speaker = "Менеджер" if u["speaker"] == 0 else "Клієнт"
                dialogue.append(f"{speaker}: {u['transcript']}")

            return "\n".join(dialogue)

        else:

            return result["results"]["channels"][0]["alternatives"][0]["transcript"]

    except Exception:

        return None


# ---------------- COMMENT ----------------

def generate_comment(dialogue):

    prompt = f"""
Проаналізуй дзвінок менеджера контакт-центру.

Напиши короткий коментар:
• коротке резюме дзвінка
• одну рекомендацію менеджеру

{dialogue}
"""

    response = client.chat.completions.create(
        model="gpt-4.1",
        temperature=0.2,
        messages=[
            {"role":"system","content":"Ти QA експерт контакт-центру"},
            {"role":"user","content":prompt}
        ]
    )

    return response.choices[0].message.content.strip()


# ---------------- SCORING ----------------

def score_call(meta):

    return {

    "Привітання":5,
    "Дружелюбне питання / Мета дзвінка":2.5,
    "Спроба продовжити розмову":5,
    "Спроба презентації":5,
    "Домовленість про наступний контакт":10,
    "Пропозиція бонусу":10,
    "Завершення":5,

    "Передзвон клієнту":10,
    "Не додумувати":5,
    "Якість мовлення":meta["speech_score"],
    "Професіоналізм":10,
    "CRM-картка":5,
    "Робота із запереченнями":10,
    "Зливання клієнта":15
    }


# ---------------- UI ----------------

st.title("🎧 QA-10: Аналіз дзвінків")

check_date = st.date_input("Дата перевірки", datetime.today())

qa_managers = [
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

    for col, idx in zip([col1, col2], [row*2+1, row*2+2]):

        with col.expander(f"📞 Дзвінок {idx}"):

            audio_url = st.text_input("Посилання на аудіо", key=f"url_{idx}")

            qa_manager = st.selectbox("QA менеджер", qa_managers, key=f"qa_{idx}")

            ret_manager = st.text_input("Менеджер RET", key=f"ret_{idx}")

            client_id = st.text_input("ID клієнта", key=f"id_{idx}")

            call_date = st.text_input("Дата дзвінка", key=f"date_{idx}")

            speech_score = st.selectbox("Якість мовлення",[2.5,0],key=f"speech_{idx}")

            calls.append({
                "url":audio_url,
                "qa_manager":qa_manager,
                "ret_manager":ret_manager,
                "client_id":client_id,
                "call_date":call_date,
                "check_date":check_date.strftime("%d-%m-%Y"),
                "speech_score":speech_score
            })


# ---------------- STATUS ----------------

status = st.empty()
progress = st.progress(0)


# ---------------- ANALYSIS ----------------

if st.button("Запустити аналіз"):

    status.info("Аналіз запущено...")

    st.session_state["results"] = []

    google_client = connect_google()

    total = len([c for c in calls if c["url"]])
    done = 0

    for i, call in enumerate(calls):

        if not call["url"]:
            continue

        status.write(f"Аналіз дзвінка {i+1}")

        transcript = transcribe_audio(call["url"])

        if not transcript:
            transcript = "Транскрипцію отримати не вдалося."

        scores = score_call(call)

        comment = generate_comment(transcript)

        write_to_google_sheet(google_client, call, scores)

        st.session_state["results"].append({
            "meta":call,
            "scores":scores,
            "comment":comment
        })

        done += 1
        progress.progress(done/total if total else 0)

    status.success("Аналіз готовий")


# ---------------- RESULTS ----------------

if "results" in st.session_state:

    for i,res in enumerate(st.session_state["results"]):

        with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):

            df = pd.DataFrame(res["scores"].items(),columns=["Критерій","Оцінка"])

            df["Оцінка"] = df["Оцінка"].apply(format_score)

            st.table(df)

            total = format_score(sum(res["scores"].values()))

            st.markdown(f"### Загальний бал: {total}")

            st.markdown("### Коментар QA")

            st.write(res["comment"])


# ---------------- EXPORT ----------------

if "results" in st.session_state and st.session_state["results"]:

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

        for i,res in enumerate(st.session_state["results"]):

            name=f"Call_{i+1}"

            meta_df=pd.DataFrame(list(res["meta"].items()),columns=["Поле","Значення"])
            meta_df.to_excel(writer,index=False,sheet_name=name)

            scores_df=pd.DataFrame(res["scores"].items(),columns=["Критерій","Оцінка"])
            scores_df["Оцінка"]=scores_df["Оцінка"].apply(format_score)

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

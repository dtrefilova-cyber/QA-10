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

"Привітання": 8,
"Дружелюбне питання / Мета дзвінка": 9,
"Спроба продовжити розмову": 10,
"Спроба презентації": 11,
"Домовленість про наступний контакт": 12,
"Пропозиція бонусу": 13,
"Завершення": 14,

"Передзвон клієнту": 17,
"Не додумувати": 18,
"Якість мовлення": 19,
"Професіоналізм": 20,
"CRM-картка": 21,
"Робота із запереченнями": 22,
"Зливання клієнта": 23
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

    cell_list = []

    for row, value in updates:
        cell_list.append(gspread.Cell(row, column, value))

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
        "punctuate": "true"
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


# ---------------- COMMENT ----------------

def generate_comment(dialogue):

    prompt = f"""
Проаналізуй дзвінок менеджера.

Напиши короткий коментар (2-3 речення):
1 коротке резюме дзвінка
1 порада менеджеру

Без списків.

{dialogue}
"""

    response = client.chat.completions.create(
        model="gpt-4.1",
        temperature=0.2,
        messages=[
            {"role": "system", "content": "Ти QA експерт контакт-центру."},
            {"role": "user", "content": prompt}
        ]
    )

    return response.choices[0].message.content.strip()


# ---------------- SCORING ----------------

def score_call(meta):

    scores = {

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

        with col.expander(f"📞 Дзвінок {idx}", expanded=False):

            audio_url = st.text_input("Посилання на аудіо", key=f"url_{idx}")

            qa_manager = st.selectbox("QA менеджер", qa_managers_list, key=f"qa_{idx}")

            ret_manager = st.text_input("Менеджер RET", key=f"ret_{idx}")

            client_id = st.text_input("ID клієнта", key=f"client_{idx}")

            call_date = st.text_input("Дата дзвінка", key=f"date_{idx}")

            speech_score = st.selectbox("Якість мовлення", [2.5,0], key=f"speech_{idx}")

            calls.append({
                "url":audio_url,
                "qa_manager":qa_manager,
                "ret_manager":ret_manager,
                "client_id":client_id,
                "call_date":call_date,
                "check_date":check_date.strftime("%d-%m-%Y"),
                "speech_score":speech_score
            })


# ---------------- ANALYSIS ----------------

if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("Запустити аналіз"):

    st.session_state["results"].clear()

    google_client = connect_google()

    for call in calls:

        if not call["url"]:
            continue

        transcript = transcribe_audio(call["url"])

        if not transcript:
            continue

        scores = score_call(call)

        comment = generate_comment(transcript)

        try:

            spreadsheet = google_client.open(call["ret_manager"])
            sheet = spreadsheet.sheet1

            write_to_google_sheet(sheet, call, scores)

        except Exception as e:
            st.warning(f"Google Sheets error: {e}")

        st.session_state["results"].append({
            "meta":call,
            "scores":scores,
            "comment":comment
        })


# ---------------- OUTPUT ----------------

for i,res in enumerate(st.session_state["results"]):

    with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):

        df = pd.DataFrame(res["scores"].items(),columns=["Критерій","Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(format_score)

        st.table(df)

        total_score = format_score(sum(res["scores"].values()))

        st.markdown(f"### Загальний бал: {total_score}")

        st.markdown("### Коментар QA")

        st.write(res["comment"])


# ---------------- EXPORT EXCEL ----------------

if st.session_state["results"]:

    xls = BytesIO()

    with pd.ExcelWriter(xls, engine="openpyxl") as writer:

        for i,res in enumerate(st.session_state["results"]):

            sheet_name=f"Call_{i+1}"

            meta_df=pd.DataFrame(list(res["meta"].items()),columns=["Поле","Значення"])
            meta_df.to_excel(writer,index=False,sheet_name=sheet_name)

            scores_df=pd.DataFrame(res["scores"].items(),columns=["Критерій","Оцінка"])
            scores_df["Оцінка"]=scores_df["Оцінка"].apply(format_score)

            scores_df.to_excel(writer,index=False,sheet_name=sheet_name,startrow=len(meta_df)+2)

            comment_df=pd.DataFrame([["Коментар",res["comment"]]],columns=["Поле","Значення"])

            comment_df.to_excel(writer,index=False,sheet_name=sheet_name,startrow=len(meta_df)+len(scores_df)+4)

    xls.seek(0)

    st.download_button(
        "📥 Завантажити результати у XLSX",
        xls,
        "qa_results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

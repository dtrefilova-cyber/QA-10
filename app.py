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
    "call_date":1,
    "qa_manager":2,
    "client_id":3,
    "check_date":4
}

CRITERIA_ROWS = {

"Привітання":5,
"Дружелюбне питання / Мета дзвінка":6,
"Спроба продовжити розмову":7,
"Спроба презентації":8,
"Пропозиція бонусу":9,
"Завершення":10,
"Домовленість про наступний контакт":11,
"Передзвон клієнту":12,
"Не додумувати":13,
"Якість мовлення":14,
"Професіоналізм":15,
"CRM-картка":16,
"Робота із запереченнями":17,
"Зливання клієнта":18
}


# ---------------- GOOGLE WRITE ----------------

def find_next_column(sheet):

    row = sheet.row_values(META_ROWS["client_id"])

    for i,v in enumerate(row,start=1):
        if v=="":
            return i

    return len(row)+1


def find_manager_sheet(client, manager_name):

    manager_name = manager_name.lower().strip()

    for file in client.openall():

        if manager_name in file.title.lower():
            return file

    return None


def write_to_google_sheet(client, meta, scores):

    spreadsheet = find_manager_sheet(client, meta["ret_manager"])

    if not spreadsheet:
        return

    sheet = spreadsheet.sheet1
    column = find_next_column(sheet)

    cells=[]

    cells.append(gspread.Cell(META_ROWS["call_date"],column,meta["call_date"]))
    cells.append(gspread.Cell(META_ROWS["qa_manager"],column,meta["qa_manager"]))
    cells.append(gspread.Cell(META_ROWS["client_id"],column,meta["client_id"]))
    cells.append(gspread.Cell(META_ROWS["check_date"],column,meta["check_date"]))

    for criterion,score in scores.items():

        if criterion in CRITERIA_ROWS:

            row=CRITERIA_ROWS[criterion]
            cells.append(gspread.Cell(row,column,float(score)))

    sheet.update_cells(cells)


# ---------------- TRANSCRIBE ----------------

def transcribe_audio(audio_url):

    if not audio_url:
        return ""

    response = requests.post(
        "https://api.deepgram.com/v1/listen",
        headers={"Authorization":f"Token {DEEPGRAM_API_KEY}"},
        params={
            "model":"nova-3",
            "language":"uk",
            "punctuate":True,
            "diarize":True,
            "utterances":True
        },
        json={"url":audio_url}
    )

    if response.status_code != 200:
        return ""

    data=response.json()

    try:

        dialogue=[]

        for u in data["results"]["utterances"]:

            speaker="Менеджер" if u["speaker"]==0 else "Клієнт"
            dialogue.append(f"{speaker}: {u['transcript']}")

        return "\n".join(dialogue)

    except:

        return data["results"]["channels"][0]["alternatives"][0]["transcript"]


# ---------------- GPT ANALYSIS ----------------

def extract_features(dialogue):

    prompt=f"""
Ти QA-аналітик контакт-центру.

Проаналізуй дзвінок менеджера.

Визнач:

1. Чи представився менеджер
2. Чи звернувся до клієнта по імені
3. Чи намагався продовжити розмову
4. Чи була презентація
5. Чи запропонував бонус
6. Чи було коректне завершення
7. Тип follow-up
8. Чи було заперечення

Поверни JSON:

{{
"manager_introduced_self":true/false,
"client_name_used":true/false,
"manager_active":true/false,
"presentation":true/false,
"bonus":true/false,
"closing":true/false,
"followup":"none / offer / day / exact_time",
"objection":true/false
}}

Дзвінок:

{dialogue}
"""

    try:

        response=client.chat.completions.create(
            model="gpt-4.1",
            temperature=0,
            messages=[
                {"role":"system","content":"Ти QA-аналітик"},
                {"role":"user","content":prompt}
            ]
        )

        text=response.choices[0].message.content

        match=re.search(r"\{[\s\S]*\}",text)

        if match:
            return json.loads(match.group())

    except:
        pass

    return {
        "manager_introduced_self":False,
        "client_name_used":False,
        "manager_active":False,
        "presentation":False,
        "bonus":False,
        "closing":False,
        "followup":"none",
        "objection":False
    }


# ---------------- SCORING ----------------

def score_call(f,meta):

    scores={}

    if f["manager_introduced_self"] and f["client_name_used"]:
        scores["Привітання"]=5
    elif f["manager_introduced_self"] or f["client_name_used"]:
        scores["Привітання"]=2.5
    else:
        scores["Привітання"]=0

    scores["Дружелюбне питання / Мета дзвінка"]=2.5

    scores["Спроба продовжити розмову"]=5 if f["manager_active"] else 0

    scores["Спроба презентації"]=5 if f["presentation"] else 0

    scores["Пропозиція бонусу"]=5 if f["bonus"] else 0

    scores["Завершення"]=5 if f["closing"] else 0

    follow=f["followup"]

    if follow=="exact_time":
        scores["Домовленість про наступний контакт"]=10
    elif follow=="day":
        scores["Домовленість про наступний контакт"]=7.5
    elif follow=="offer":
        scores["Домовленість про наступний контакт"]=5
    else:
        scores["Домовленість про наступний контакт"]=0

    repeat=meta["repeat_call"]

    if repeat=="так, був протягом години":
        scores["Передзвон клієнту"]=10
    elif repeat=="так, був протягом 3 годин":
        scores["Передзвон клієнту"]=5
    else:
        scores["Передзвон клієнту"]=0 if follow!="none" else 10

    scores["Не додумувати"]=5
    scores["Якість мовлення"]=meta["speech_score"]

    scores["Професіоналізм"]=5 if meta["bonus_check"]=="помилково нараховано" else 10

    scores["CRM-картка"]=5 if meta["manager_comment"] else 0

    scores["Робота із запереченнями"]=10 if not f["objection"] else 5

    scores["Зливання клієнта"]=15

    return scores


def format_score(x):
    return f"{float(x):.1f}"
    # ---------------- UI ----------------

st.title("🎧 QA-10")

check_date = st.date_input("Дата перевірки", datetime.today())

qa_list = [
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

    c1, c2 = st.columns(2)

    for col, idx in zip([c1, c2], [row*2+1, row*2+2]):

        with col.expander(f"📞 Дзвінок {idx}"):

            url = st.text_input("Посилання на дзвінок", key=f"url_{idx}")
            qa = st.selectbox("QA менеджер", qa_list, key=f"qa_{idx}")
            ret = st.text_input("Менеджер RET", key=f"ret_{idx}")
            cid = st.text_input("ID клієнта", key=f"id_{idx}")
            date = st.text_input("Дата дзвінка", key=f"date_{idx}")

            bonus = st.selectbox(
                "Бонус",
                ["правильно нараховано", "помилково нараховано", "не потрібно"],
                key=f"bonus_{idx}"
            )

            repeat = st.selectbox(
                "Повторний дзвінок",
                ["так, був протягом години", "так, був протягом 3 годин", "ні, не було"],
                key=f"repeat_{idx}"
            )

            comment = st.text_area("Коментар з картки", key=f"comment_{idx}")

            speech = st.selectbox("Оцінка мови", [2.5, 0], key=f"speech_{idx}")

            calls.append({
                "url": url,
                "qa_manager": qa,
                "ret_manager": ret,
                "client_id": cid,
                "call_date": date,
                "check_date": check_date.strftime("%d-%m-%Y"),
                "bonus_check": bonus,
                "repeat_call": repeat,
                "manager_comment": comment,
                "speech_score": speech
            })


if "results" not in st.session_state:
    st.session_state["results"] = []

status = st.empty()


# ---------------- ANALYSIS ----------------

if st.button("Запустити аналіз"):

    st.session_state["results"].clear()

    google_client = connect_google()

    for i, call in enumerate(calls):

        if not call["url"]:
            continue

        status.write(f"⏳ Аналіз дзвінка {i+1}")

        transcript = transcribe_audio(call["url"])

        features = extract_features(transcript)

        scores = score_call(features, call)

        comment = generate_comment(transcript)

        write_to_google_sheet(google_client, call, scores)

        st.session_state["results"].append({
            "meta": call,
            "scores": scores,
            "comment": comment
        })

    status.success("Аналіз завершено")


# ---------------- RESULTS ----------------

for i, res in enumerate(st.session_state["results"]):

    with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):

        df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(format_score)

        st.table(df)

        total_score = f"{sum(res['scores'].values()):.1f}"

        st.markdown(f"**Загальний бал:** {total_score}")

        st.markdown("### Коментар")
        st.write(res["comment"])


# ---------------- EXPORT ----------------

if st.session_state["results"]:

    xls = BytesIO()

    with pd.ExcelWriter(xls, engine="openpyxl") as writer:

        for i, res in enumerate(st.session_state["results"]):

            sheet = f"Call_{i+1}"

            meta_df = pd.DataFrame(
                list(res["meta"].items()),
                columns=["Поле", "Значення"]
            )

            meta_df.to_excel(
                writer,
                index=False,
                sheet_name=sheet
            )

            scores_df = pd.DataFrame(
                res["scores"].items(),
                columns=["Критерій", "Оцінка"]
            )

            scores_df["Оцінка"] = scores_df["Оцінка"].apply(format_score)

            scores_df.to_excel(
                writer,
                index=False,
                sheet_name=sheet,
                startrow=len(meta_df) + 2
            )

            comment_df = pd.DataFrame(
                [["Коментар", res["comment"]]],
                columns=["Поле", "Значення"]
            )

            comment_df.to_excel(
                writer,
                index=False,
                sheet_name=sheet,
                startrow=len(meta_df) + len(scores_df) + 4
            )

    xls.seek(0)

    st.download_button(
        "📥 Завантажити результати XLSX",
        xls,
        "qa_results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

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

    response=requests.post(
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

    if response.status_code!=200:
        return ""

    data=response.json()

    dialogue=[]

    try:

        for u in data["results"]["utterances"]:

            speaker="Менеджер" if u["speaker"]==0 else "Клієнт"
            dialogue.append(f"{speaker}: {u['transcript']}")

        return "\n".join(dialogue)

    except:

        return data["results"]["channels"][0]["alternatives"][0]["transcript"]


# ---------------- GPT QA ----------------

def analyze_call(dialogue,meta):

    prompt=f"""
Ти QA-аналітик контакт-центру.

Поверни ТІЛЬКИ JSON.

{{
"Привітання":number,
"Дружелюбне питання / Мета дзвінка":number,
"Спроба продовжити розмову":number,
"Спроба презентації":number,
"Домовленість про наступний контакт":number,
"Пропозиція бонусу":number,
"Завершення":number,
"Передзвон клієнту":number,
"Не додумувати":number,
"Якість мовлення":number,
"Професіоналізм":number,
"CRM-картка":number,
"Робота із запереченнями":number,
"Зливання клієнта":number,
"comment":"string"
}}

Дзвінок:

{dialogue}

Коментар менеджера:
{meta["manager_comment"]}
"""

    try:

        response = client.chat.completions.create(
            model="gpt-4.1",
            temperature=0,
            messages=[
                {"role":"system","content":"Ти QA-аналітик контакт-центру"},
                {"role":"user","content":prompt}
            ]
        )

        text=response.choices[0].message.content.strip()

        text=text.replace("```json","").replace("```","")

        match=re.search(r"\{[\s\S]*\}",text)

        if match:
            result=json.loads(match.group())
        else:
            raise ValueError

    except:

        result={}

    template={
    "Привітання":0,
    "Дружелюбне питання / Мета дзвінка":0,
    "Спроба продовжити розмову":0,
    "Спроба презентації":0,
    "Домовленість про наступний контакт":0,
    "Пропозиція бонусу":0,
    "Завершення":0,
    "Передзвон клієнту":0,
    "Не додумувати":0,
    "Якість мовлення":0,
    "Професіоналізм":0,
    "CRM-картка":0,
    "Робота із запереченнями":0,
    "Зливання клієнта":0,
    "comment":""
    }

    for k in template:
        if k not in result:
            result[k]=template[k]

    return result


def format_score(x):
    return f"{float(x):.1f}"


# ---------------- UI ----------------

st.title("🎧 QA-10")

check_date=st.date_input("Дата перевірки",datetime.today())

qa_list=[
"Аліна","Дар'я","Надя","Настя",
"Владимира","Діана","Руслана","Олексій"
]

calls=[]

for row in range(5):

    c1,c2=st.columns(2)

    for col,idx in zip([c1,c2],[row*2+1,row*2+2]):

        with col.expander(f"📞 Дзвінок {idx}"):

            url=st.text_input("Посилання на дзвінок",key=f"url_{idx}")
            qa=st.selectbox("QA менеджер",qa_list,key=f"qa_{idx}")
            ret=st.text_input("Менеджер RET",key=f"ret_{idx}")
            cid=st.text_input("ID клієнта",key=f"id_{idx}")
            date=st.text_input("Дата дзвінка",key=f"date_{idx}")

            repeat=st.selectbox(
            "Повторний дзвінок",
            ["так, був протягом години","так, був протягом 3 годин","ні, не було"],
            key=f"repeat_{idx}"
            )

            comment=st.text_area("Коментар з картки",key=f"comment_{idx}")

            speech=st.selectbox("Оцінка мови",[2.5,0],key=f"speech_{idx}")

            calls.append({
            "url":url,
            "qa_manager":qa,
            "ret_manager":ret,
            "client_id":cid,
            "call_date":date,
            "check_date":check_date.strftime("%d-%m-%Y"),
            "repeat_call":repeat,
            "manager_comment":comment,
            "speech_score":speech
            })


if "results" not in st.session_state:
    st.session_state["results"]=[]

status=st.empty()


# ---------------- ANALYSIS ----------------

if st.button("Запустити аналіз"):

    st.session_state["results"].clear()

    google_client=connect_google()

    for i,call in enumerate(calls):

        if not call["url"]:
            continue

        status.write(f"Аналіз дзвінка {i+1}")

        transcript=transcribe_audio(call["url"])

        result=analyze_call(transcript,call)

        scores={k:v for k,v in result.items() if k!="comment"}

        comment=result.get("comment","")

        write_to_google_sheet(google_client,call,scores)

        st.session_state["results"].append({
        "meta":call,
        "scores":scores,
        "comment":comment
        })

    status.success("Аналіз завершено")


# ---------------- RESULTS ----------------

for i,res in enumerate(st.session_state["results"]):

    with st.expander(f"Результат дзвінка {i+1}",expanded=True):

        df=pd.DataFrame(res["scores"].items(),columns=["Критерій","Оцінка"])
        df["Оцінка"]=df["Оцінка"].apply(format_score)

        st.table(df)

        total_score=f"{sum(res['scores'].values()):.1f}"

        st.markdown(f"Загальний бал: **{total_score}**")

        st.markdown("Коментар")
        st.write(res["comment"])


# ---------------- EXPORT ----------------

if st.session_state["results"]:

    xls=BytesIO()

    with pd.ExcelWriter(xls,engine="openpyxl") as writer:

        for i,res in enumerate(st.session_state["results"]):

            sheet=f"Call_{i+1}"

            meta_df=pd.DataFrame(list(res["meta"].items()),columns=["Поле","Значення"])
            meta_df.to_excel(writer,index=False,sheet_name=sheet)

            scores_df=pd.DataFrame(res["scores"].items(),columns=["Критерій","Оцінка"])
            scores_df["Оцінка"]=scores_df["Оцінка"].apply(format_score)

            scores_df.to_excel(writer,index=False,sheet_name=sheet,startrow=len(meta_df)+2)

    xls.seek(0)

    st.download_button(
    "Завантажити XLSX",
    xls,
    "qa_results.xlsx"
    )

import streamlit as st
import pandas as pd
import requests
import json
import re
from google_sheets import connect_google, write_to_google_sheet
from io import BytesIO
from datetime import datetime
from openai import OpenAI
from prompts import get_full_analysis_prompt

# ================= CONFIG =================
DEEPGRAM_API_KEY = st.secrets["DEEPGRAM_API_KEY"]
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]
LOG_SHEET_ID = st.secrets["LOG_SHEET_ID"]

client = OpenAI(api_key=OPENAI_API_KEY)

st.title("🎧 QA-10: Аналіз дзвінків")

check_date = st.date_input("Дата перевірки", datetime.today())

qa_managers_list = [
    "Дар'я", "Надя", "Настя", "Владимира", "Діана", "Руслана", "Олексій"
]

# ================= INPUT =================
calls = []
for row in range(5):
    col1, col2 = st.columns(2)
    for col, idx in zip([col1, col2], [row * 2 + 1, row * 2 + 2]):
        with col.expander(f"📞 Дзвінок {idx}"):

            audio_url = st.text_input("Посилання", key=f"url_{idx}")
            qa_manager = st.selectbox("QA", qa_managers_list, key=f"qa_{idx}")
            ret_manager = st.text_input("Менеджер", key=f"ret_{idx}")
            client_id = st.text_input("ID", key=f"client_{idx}")
            call_date = st.text_input("Дата", key=f"date_{idx}")

            bonus_check = st.selectbox(
                "Бонус",
                ["правильно нараховано", "помилково нараховано", "не потрібно"],
                key=f"bonus_{idx}
            )

            repeat_call = st.selectbox(
                "Передзвон",
                ["так, був протягом години", "так, був протягом 2 годин", "ні, не було"],
                key=f"repeat_{idx}"
            )

            manager_comment = st.text_area("Коментар", key=f"comment_{idx}")
            speech_score = st.selectbox("Мовлення", [2.5, 0], key=f"speech_{idx}")

            calls.append({
                "url": audio_url.strip(),
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

# ================= TRANSCRIPTION =================
def clean_transcript(text):
    replacements = {
        "вагас": "Vegas",
        "вегас": "Vegas",
        "відпрограма": "віп програма"
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text


def transcribe_audio(url):
    if not url:
        return None

    try:
        r = requests.post(
            "https://api.deepgram.com/v1/listen",
            headers={"Authorization": f"Token {DEEPGRAM_API_KEY}"},
            params={"model": "nova-3", "language": "uk"},
            json={"url": url}
        )

        data = r.json()
        text = data["results"]["channels"][0]["alternatives"][0]["transcript"]

        return clean_transcript(text)

    except:
        return None


def extract_segments(dialogue):
    lines = dialogue.split("\n")
    return "\n".join(lines[:5]), "\n".join(lines[5:-5]), "\n".join(lines[-5:])


# ================= GPT =================
def extract_features(dialogue):
    intro, middle, ending = extract_segments(dialogue)
    prompt = get_full_analysis_prompt(intro, middle, ending)

    try:
        res = client.chat.completions.create(
            model="gpt-5.4",
            temperature=0,
            messages=[
                {"role": "system", "content": "JSON only"},
                {"role": "user", "content": prompt}
            ]
        )

        text = res.choices[0].message.content
        match = re.search(r"\{[\s\S]*\}", text)
        return json.loads(match.group()) if match else {}

    except:
        return {}


# ================= SCORING =================
def score_call(f, meta):
    s = {}

    elements = sum([
        f.get("manager_name_present"),
        f.get("manager_position_present"),
        f.get("company_present"),
        f.get("client_name_used"),
        f.get("purpose_present")
    ])

    s["Встановлення контакту"] = 7.5 if elements >= 4 else 5 if elements == 3 else 2.5 if elements == 2 else 0

    s["Спроба презентації"] = f.get("presentation_score", 0)

    fup = f.get("followup_type", "none")
    s["Домовленість про наступний контакт"] = 5 if fup == "exact_time" else 2.5 if fup == "offer" else 0

    cond = len(set(f.get("bonus_conditions", [])))
    s["Пропозиція бонусу"] = 10 if cond >= 2 else 5 if cond == 1 else 0

    s["Завершення розмови"] = 5 if f.get("has_farewell") else 0

    s["Передзвон клієнту"] = 15 if meta["repeat_call"] == "так, був протягом години" else 10 if meta["repeat_call"] == "так, був протягом 2 годин" else 0

    s["Не додумувати"] = 5
    s["Якість мовлення"] = meta["speech_score"]
    s["Професіоналізм"] = 5 if meta["bonus_check"] == "помилково нараховано" else 10

    comment = meta["manager_comment"]
    s["Оформлення картки"] = 0 if not comment else 2.5 if len(comment.split()) < 4 else 5

    lvl = f.get("continuation_level", "none")
    s["Утримання клієнта"] = 20 if lvl == "strong" else 15 if lvl == "weak" else 10

    s["Робота із запереченнями"] = 10 if not f.get("objection_detected") else (10 if lvl == "strong" else 5 if lvl == "weak" else 0)

    return s


# ================= RUN =================
if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("🚀 Запустити аналіз"):

    st.session_state["results"].clear()
    google_client = connect_google()

    for call in calls:
        if not call["url"]:
            continue

        transcript = transcribe_audio(call["url"])
        if not transcript:
            continue

        features = extract_features(transcript)
        scores = score_call(features, call)

        st.session_state["results"].append({"scores": scores})


# ================= OUTPUT =================
for res in st.session_state["results"]:
    df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
    st.table(df)
	# ================= COMMENT =================
def generate_qa_comment(scores, features):
    comments = []

    if scores["Встановлення контакту"] < 7.5:
        comments.append("Встановлення контакту — не всі елементи виконані")

    if scores["Спроба презентації"] == 0:
        comments.append("Спроба презентації — відсутня презентація")

    if scores["Домовленість про наступний контакт"] < 5:
        comments.append("Домовленість про наступний контакт — немає точного часу")

    if scores["Пропозиція бонусу"] < 10:
        comments.append("Пропозиція бонусу — недостатньо умов")

    if scores["Завершення розмови"] < 5:
        comments.append("Завершення розмови — немає коректного прощання")

    if scores["Утримання клієнта"] < 20:
        comments.append("Утримання клієнта — слабка робота з клієнтом")

    if not comments:
        return "Усі критерії виконані на максимум"

    return "\n".join([f"- {c}" for c in comments])


# ================= RUN =================
if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("🚀 Запустити аналіз", type="primary"):

    st.session_state["results"].clear()

    google_client = None
    try:
        google_client = connect_google()
    except Exception as e:
        st.error(f"Google error: {e}")

    for i, call in enumerate(calls):

        if not call["url"]:
            continue

        st.write(f"Обробка дзвінка {i+1}")

        transcript = transcribe_audio(call["url"])
        if not transcript:
            st.warning("Немає транскрипції")
            continue

        features = extract_features(transcript)
        scores = score_call(features, call)
        comment = generate_qa_comment(scores, features)

        # ================= GOOGLE SHEETS =================
        if google_client:
            try:
                # основна таблиця
                sheet = google_client.open(call["ret_manager"]).sheet1
                write_to_google_sheet(sheet, call, scores)

                # запис коментаря (A/B з 20 рядка)
                start_row = 20
                existing_ids = sheet.col_values(1)[start_row-1:]
                next_row = start_row + len(existing_ids)

                sheet.update(f"A{next_row}", call["client_id"])
                sheet.update(f"B{next_row}", comment)

                # лог таблиця
                log_spreadsheet = google_client.open_by_key(LOG_SHEET_ID)
                log_sheet = log_spreadsheet.sheet1

                if transcript and len(transcript.strip()) > 10:
                    log_sheet.append_row([
                        call["check_date"],
                        call["client_id"],
                        call["ret_manager"],
                        transcript,
                        comment,
                        sum(scores.values())
                    ])

            except Exception as e:
                st.error(f"Google Sheets error: {e}")

        st.session_state["results"].append({
            "scores": scores,
            "comment": comment
        })


# ================= OUTPUT =================
for i, res in enumerate(st.session_state["results"]):

    with st.expander(f"📊 Дзвінок {i+1}", expanded=True):

        df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
        st.table(df)

        total = sum(res["scores"].values())
        st.success(f"Загальний бал: {total:.1f}")

        st.markdown("### Коментар QA")
        st.write(res["comment"])


# ================= EXPORT =================
if st.session_state["results"]:

    xls = BytesIO()

    with pd.ExcelWriter(xls, engine="openpyxl") as writer:
        for i, res in enumerate(st.session_state["results"]):
            df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
            df.to_excel(writer, sheet_name=f"Call_{i+1}", index=False)

    xls.seek(0)

    st.download_button(
        label="📥 Завантажити Excel",
        data=xls,
        file_name="qa_results.xlsx"
    )

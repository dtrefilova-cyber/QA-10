import streamlit as st
import pandas as pd
import requests
import json
import re
from io import BytesIO
from datetime import datetime
from openai import OpenAI

from prompts import get_full_analysis_prompt
from google_sheets import connect_google, write_to_google_sheet

DEEPGRAM_API_KEY = st.secrets["DEEPGRAM_API_KEY"]
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]

client = OpenAI(api_key=OPENAI_API_KEY)

st.title("🎧 QA-10: Аналіз дзвінків")

check_date = st.date_input("Дата перевірки", datetime.today())

qa_managers_list = [
    "Дар'я", "Надя", "Настя", "Владимира", "Діана", "Руслана", "Олексій"
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
            
            bonus_check = st.selectbox(
                "Бонус", 
                ["правильно нараховано", "помилково нараховано", "не потрібно"], 
                key=f"bonus_{idx}"
            )
            repeat_call = st.selectbox(
                "Повторний дзвінок",
                ["так, був протягом години", "так, був протягом 3 годин", "ні, не було"],
                key=f"repeat_{idx}"
            )
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


# ====================== TRANSCRIPTION ======================
def transcribe_audio(audio_url):
    if not audio_url:
        return None

    url = "https://api.deepgram.com/v1/listen"
    params = {
        "model": "nova-3",
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

    if "results" not in result or "utterances" not in result["results"]:
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


def extract_segments(dialogue):
    lines = dialogue.split("\n")
    intro = "\n".join(lines[:5])
    middle = "\n".join(lines[5:-5]) if len(lines) > 10 else "\n".join(lines[5:])
    ending = "\n".join(lines[-5:]) if len(lines) > 5 else ""
    return intro, middle, ending


# ====================== GPT ======================
def extract_features(dialogue):
    intro, middle, ending = extract_segments(dialogue)
    prompt = get_full_analysis_prompt(intro, middle, ending)

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            temperature=0,
            messages=[
                {"role": "system", "content": "Відповідай тільки валідним JSON"},
                {"role": "user", "content": prompt}
            ]
        )

        text = response.choices[0].message.content.strip()
        match = re.search(r"\{[\s\S]*\}", text)
        data = json.loads(match.group()) if match else {}

    except Exception as e:
        st.warning(f"Помилка GPT: {e}")
        data = {}

    return data


# ====================== SCORING ======================
def score_call(features, meta):
    scores = {
        "Встановлення контакту": features.get("contact_score", 0),
        "Спроба презентації": features.get("presentation_score", 0),
        "Домовленість про наступний контакт": features.get("followup_score", 0),
        "Пропозиція бонусу": features.get("bonus_score", 0),
        "Завершення розмови": features.get("closing_score", 0),
        "Передзвон клієнту": features.get("callback_score", 0),
        "Не додумувати": features.get("no_assumption_score", 0),
        "Якість мовлення": meta["speech_score"],
        "Професіоналізм": features.get("professionalism_score", 0),
        "Оформлення картки": features.get("crm_score", 0),
        "Робота із запереченнями": features.get("objection_score", 0),
        "Утримання клієнта": features.get("retention_score", 0),
    }

    return scores


# ====================== COMMENT ======================
def generate_comment(dialogue):
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            temperature=0.3,
            messages=[{
                "role": "user",
                "content": f"Підсумуй дзвінок у 1-2 реченнях. Сильна сторона + 1 рекомендація.\n{dialogue}"
            }]
        )
        return response.choices[0].message.content
    except:
        return "Не вдалося згенерувати коментар."


# ====================== RUN ======================
if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("🚀 Запустити аналіз", type="primary"):
    st.session_state["results"].clear()

    try:
        google_client = connect_google()
    except Exception as e:
        st.warning(f"Google Sheets помилка: {e}")
        google_client = None

    for i, call in enumerate(calls):
        if not call["url"].strip():
            continue

        with st.spinner(f"Обробка дзвінка {i+1}..."):
            transcript = transcribe_audio(call["url"])

            if not transcript:
                st.error(f"Помилка транскрипції {i+1}")
                continue

            features = extract_features(transcript)
            scores = score_call(features, call)
            comment = generate_comment(transcript)

            if google_client:
                try:
                    spreadsheet = google_client.open(call["ret_manager"])
                    sheet = spreadsheet.sheet1
                    write_to_google_sheet(sheet, call, scores)
                except Exception as e:
                    st.warning(f"Помилка Google Sheets: {e}")

            st.session_state["results"].append({
                "meta": call,
                "scores": scores,
                "comment": comment,
                "features": features
            })


# ====================== OUTPUT ======================
for i, res in enumerate(st.session_state["results"]):
    with st.expander(f"📊 Дзвінок {i+1}", expanded=True):
        df = pd.DataFrame(list(res["scores"].items()), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
        st.table(df)

        total_score = res["features"].get("total_score", sum(res["scores"].values()))
        st.success(f"Загальний бал: {total_score:.1f}")

        st.markdown("### Коментар")
        st.write(res["comment"])


# ====================== EXPORT ======================
if st.session_state["results"]:
    xls = BytesIO()

    with pd.ExcelWriter(xls, engine="openpyxl") as writer:
        for i, res in enumerate(st.session_state["results"]):
            sheet_name = f"Call_{i+1}"

            meta_df = pd.DataFrame(list(res["meta"].items()), columns=["Поле", "Значення"])
            meta_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)

            scores_df = pd.DataFrame(list(res["scores"].items()), columns=["Критерій", "Оцінка"])
            scores_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=len(meta_df) + 2)

            comment_df = pd.DataFrame([["Коментар", res["comment"]]], columns=["Поле", "Значення"])
            comment_df.to_excel(writer, index=False, sheet_name=sheet_name,
                                startrow=len(meta_df) + len(scores_df) + 4)

    xls.seek(0)

    st.download_button(
        label="📥 Завантажити XLSX",
        data=xls,
        file_name="qa_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

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
                {"role": "system", "content": "Поверни тільки JSON"},
                {"role": "user", "content": prompt}
            ]
        )
        text = response.choices[0].message.content.strip()
        match = re.search(r"\{[\s\S]*\}", text)
        features = json.loads(match.group()) if match else {}
    except Exception as e:
        st.warning(f"Помилка GPT: {e}")
        features = {}

    defaults = {
        "manager_introduced_self": False,
        "client_name_used": False,
        "presentation_score": 0,
        "bonus_offered": False,
        "bonus_conditions_count": 0,
        "followup_type": "none",
        "objection_detected": False,
        "conversation_continuation_score": 0
    }
    for k, v in defaults.items():
        features.setdefault(k, v)
    
    features["raw_text"] = dialogue.lower()
    return features


# ====================== SCORING ======================
def score_call(features, meta):
    scores = {}
    raw = features.get("raw_text", "")

    # CONTACT - Встановлення контакту
    has_name = features["manager_introduced_self"]
    has_client = features["client_name_used"]
    has_company = any(w in raw for w in ["компан", "казино", "служба підтримки"])
    has_purpose = any(w in raw for w in ["телефоную", "дзвоню", "звертаюсь"])
    has_friendly = any(w in raw for w in ["як справ", "зручно говорити"])

    elements = sum([
        has_name,
        has_client,
        has_company,
        (has_purpose or has_friendly)
    ])
    scores["Встановлення контакту"] = [0, 0, 2.5, 5, 7.5][elements]

    # FOLLOWUP
    f = features["followup_type"]
    scores["Домовленість про наступний контакт"] = 5 if f == "exact_time" else 2.5 if f == "offer" else 0

    # BONUS
    if not features["bonus_offered"]:
        score = 0
    elif features["bonus_conditions_count"] >= 1:
        score = 10
    else:
        score = 5
    scores["Пропозиція бонусу"] = score

    # CLOSING
    farewell_words = [
        "до побачення", "гарного дня", "всього доброго",
        "дякую", "до зв'язку", "на все добре"
    ]
    has_farewell = any(w in raw for w in farewell_words)
    scores["Завершення розмови"] = 5 if has_farewell else 0

    # CALLBACK
    if meta["repeat_call"] == "так, був протягом години":
        scores["Передзвон клієнту"] = 15
    elif meta["repeat_call"] == "так, був протягом 3 годин":
        scores["Передзвон клієнту"] = 10
    else:
        scores["Передзвон клієнту"] = 0

    # NO ASSUMPTION
    scores["Не додумувати"] = 0 if "давайте потім" in raw else 5

    # SPEECH
    scores["Якість мовлення"] = meta["speech_score"]

    # PROFESSIONAL
    scores["Професіоналізм"] = 5 if meta["bonus_check"] == "помилково нараховано" else 10

    # CRM
    scores["Оформлення картки"] = 5 if meta["manager_comment"] else 0

    # OBJECTION
    if not features["objection_detected"]:
        scores["Робота із запереченнями"] = 10
    else:
        cont = features.get("conversation_continuation_score", 0)
        if cont == 5:
            scores["Робота із запереченнями"] = 10
        elif cont == 2.5:
            scores["Робота із запереченнями"] = 5
        else:
            scores["Робота із запереченнями"] = 0

    # RETENTION (Утримання клієнта)
    cont = features.get("conversation_continuation_score", 0)
    if cont == 5:
        score = 20
    elif cont == 2.5:
        score = 15
    else:
        score = 10
    scores["Утримання клієнта"] = score

    return scores


# ====================== COMMENT ======================
def generate_comment(dialogue):
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            temperature=0.3,
            messages=[{
                "role": "user",
                "content": f"Підсумуй дзвінок у 1-2 реченнях. Вкажи сильну сторону менеджера і одну рекомендацію.\n{dialogue}"
            }]
        )
        return response.choices[0].message.content
    except:
        return "Не вдалося згенерувати коментар."


def explain_scores(dialogue, scores):
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            temperature=0,
            messages=[{
                "role": "user",
                "content": f"""
Поясни коротко причину оцінки по кожному критерію.
Оцінки:
{scores}
Дзвінок:
{dialogue}
Формат:
Критерій: причина
"""
            }]
        )
        return response.choices[0].message.content
    except:
        return ""


# ====================== RUN ======================
if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("🚀 Запустити аналіз", type="primary"):
    st.session_state["results"].clear()

    # Google Sheets
    try:
        google_client = connect_google()
    except Exception as e:
        st.warning(f"Не вдалося підключитись до Google Sheets: {e}")
        google_client = None

    for i, call in enumerate(calls):
        if not call["url"].strip():
            continue

        with st.spinner(f"Обробка дзвінка {i+1}..."):
            transcript = transcribe_audio(call["url"])
            if not transcript:
                st.error(f"Не вдалося транскрибувати дзвінок {i+1}")
                continue

            features = extract_features(transcript)
            scores = score_call(features, call)
            comment = generate_comment(transcript)
            explanation = explain_scores(transcript, scores)

            # Google Sheets запис
            if google_client:
                try:
                    spreadsheet = google_client.open(call["ret_manager"])
                    sheet = spreadsheet.sheet1
                    write_to_google_sheet(sheet, call, scores)
                except Exception as e:
                    st.warning(f"Помилка запису в Google Sheets: {e}")

            st.session_state["results"].append({
                "meta": call,
                "scores": scores,
                "comment": comment,
                "explanation": explanation
            })


# ====================== OUTPUT ======================
for i, res in enumerate(st.session_state["results"]):
    with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):
        df = pd.DataFrame(list(res["scores"].items()), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
        st.table(df)

        total_score = sum(res["scores"].values())
        st.success(f"Загальний бал: {total_score:.1f}")

        st.markdown("### Коментар QA")
        st.write(res["comment"])
        st.markdown("### Пояснення оцінки")
        st.write(res["explanation"])


# ====================== EXPORT ======================
if st.session_state["results"]:
    xls = BytesIO()
    with pd.ExcelWriter(xls, engine="openpyxl") as writer:
        for i, res in enumerate(st.session_state["results"]):
            sheet_name = f"Call_{i+1}"
            meta_df = pd.DataFrame(list(res["meta"].items()), columns=["Поле", "Значення"])
            meta_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)

            scores_df = pd.DataFrame(list(res["scores"].items()), columns=["Критерій", "Оцінка"])
            scores_df["Оцінка"] = scores_df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
            scores_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=len(meta_df) + 2)

            comment_df = pd.DataFrame(
                [["Коментар", res["comment"]]],
                columns=["Поле", "Значення"]
            )
            comment_df.to_excel(
                writer,
                index=False,
                sheet_name=sheet_name,
                startrow=len(meta_df) + len(scores_df) + 4
            )

    xls.seek(0)
    st.download_button(
        label="📥 Завантажити результати у XLSX",
        data=xls,
        file_name="qa_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

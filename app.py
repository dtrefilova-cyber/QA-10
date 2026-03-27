import streamlit as st
import pandas as pd
import requests
import json
import re
from io import BytesIO
from datetime import datetime
from openai import OpenAI

# Імпорт промпту з окремого файлу
from prompts import get_full_analysis_prompt

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


# ====================== FEATURE EXTRACTION ======================
def extract_features(dialogue):
    intro, middle, ending = extract_segments(dialogue)

    # Отримуємо промпт з окремого файлу
    prompt = get_full_analysis_prompt(intro, middle, ending)

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            temperature=0,
            messages=[
                {"role": "system", "content": "Відповідай виключно валідним JSON. Не додавай жодного тексту крім JSON."},
                {"role": "user", "content": prompt}
            ]
        )
        text = response.choices[0].message.content.strip()
        match = re.search(r"\{[\s\S]*\}", text)
        
        features = json.loads(match.group()) if match else {}
    except Exception as e:
        st.warning(f"Помилка при аналізі дзвінка: {e}")
        features = {}

    # Дефолтні значення
    defaults = {
        "manager_introduced_self": False,
        "client_name_used": False,
        "presentation_score": 0,
        "bonus_offered": False,
        "bonus_conditions_count": 0,
        "client_busy_or_rude": False,
        "client_hung_up": False,
        "manager_active": True,
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
    
    # Fallback на ключові слова (сарказм, роздратування, матюки)
    raw = features.get("raw_text", "").lower()
    rude_indicators = ["не цікаво", "не хочу", "ага зрозуміло", "розрізся", "ти що", "ха-ха", "дура", "лохи", 
                       "бля", "хуй", "нахуй", "сука", "відчеп", "зайоб", "пизд"]
    if any(ind in raw for ind in rude_indicators):
        features["client_busy_or_rude"] = True
        features["presentation_score"] = 5
        features["bonus_offered"] = True
        features["followup_type"] = "exact_time"

    is_critical = features.get("client_busy_or_rude", False) or features.get("client_hung_up", False)

    # Привітання
    if features["manager_introduced_self"] and features["client_name_used"]:
        scores["Привітання"] = 5
    elif features["manager_introduced_self"] or features["client_name_used"]:
        scores["Привітання"] = 2.5
    else:
        scores["Привітання"] = 0

    scores["Дружелюбне питання / Мета дзвінка"] = 2.5
    
    # Спроба продовжити розмову
    if is_critical:
        scores["Спроба продовжити розмову"] = 5
    else:
        scores["Спроба продовжити розмову"] = features.get("conversation_continuation_score", 0)

    # Спроба презентації
    if is_critical:
        scores["Спроба презентації"] = 5
    else:
        presentation_score = features.get("presentation_score", 0)
        if presentation_score > 0:
            scores["Спроба презентації"] = presentation_score
        else:
            text = features.get("raw_text", "")
            pres_kw = ["слот", "бонус", "турнір", "акці", "фріспін", "активн", "надішл", "пошт", "email"]
            scores["Спроба презентації"] = 5 if any(k in text for k in pres_kw) else 0

    # Домовленість про наступний контакт
    followup = features.get("followup_type", "none")
    if is_critical or followup == "exact_time":
        scores["Домовленість про наступний контакт"] = 10
    elif followup == "day":
        scores["Домовленість про наступний контакт"] = 7.5
    elif followup == "offer":
        scores["Домовленість про наступний контакт"] = 5
    else:
        scores["Домовленість про наступний контакт"] = 0

    # Пропозиція бонусу
    if is_critical:
        scores["Пропозиція бонусу"] = 10
    else:
        bc = features["bonus_conditions_count"]
        if not features["bonus_offered"]:
            scores["Пропозиція бонусу"] = 0
        elif bc == 0:
            scores["Пропозиція бонусу"] = 5
        elif bc == 1:
            scores["Пропозиція бонусу"] = 7.5
        else:
            scores["Пропозиція бонусу"] = 10

    # Завершення - завжди 5, якщо менеджер попрощався
    # Перевіряємо наявність прощання в тексті
    raw_text = features.get("raw_text", "")
    has_farewell = any(word in raw_text for word in ["до побачення", "бувайте", "гарного дня", "всього доброго", "до зв'язку"])
    scores["Завершення"] = 5 if has_farewell else 0

    # Передзвон клієнту
    repeat = meta["repeat_call"]
    if repeat == "так, був протягом години":
        scores["Передзвон клієнту"] = 10
    elif repeat == "так, був протягом 3 годин":
        scores["Передзвон клієнту"] = 5
    else:
        scores["Передзвон клієнту"] = 0

    # Інші критерії
    scores["Не додумувати"] = 5
    scores["Якість мовлення"] = meta["speech_score"]
    scores["Професіоналізм"] = 5 if meta["bonus_check"] == "помилково нараховано" else 10
    scores["CRM-картка"] = 5 if meta.get("manager_comment") else 0
    
    # Робота із запереченнями
    if features["objection_detected"]:
        cont_score = features.get("conversation_continuation_score", 0)
        if cont_score == 5:
            scores["Робота із запереченнями"] = 10
        elif cont_score == 2.5:
            scores["Робота із запереченнями"] = 7.5
        else:
            scores["Робота із запереченнями"] = 0
    else:
        # Якщо заперечень не було - ставимо 10 балів
        scores["Робота із запереченнями"] = 10
    
    scores["Зливання клієнта"] = 15 if (is_critical or features["manager_active"]) else 10

    return scores


# ====================== GENERATE COMMENT ======================
def generate_comment(dialogue):
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            temperature=0.3,
            messages=[{"role": "user", "content": f"Підсумуй дзвінок у 1-2 реченнях. Зазнач сильну сторону менеджера та одну ключову рекомендацію.\n{dialogue}"}]
        )
        return response.choices[0].message.content
    except:
        return "Не вдалося згенерувати коментар."


# ====================== ЗАПУСК АНАЛІЗУ ======================
if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("🚀 Запустити аналіз", type="primary"):
    st.session_state["results"].clear()
    
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

            st.session_state["results"].append({
                "meta": call,
                "scores": scores,
                "comment": comment
            })

# ====================== ВИВІД РЕЗУЛЬТАТІВ ======================
for i, res in enumerate(st.session_state["results"]):
    with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):
        df = pd.DataFrame(list(res["scores"].items()), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
        st.table(df)
        
        total_score = sum(res["scores"].values())
        st.success(f"**Загальний бал: {total_score:.1f}**")
        
        st.markdown("### Коментар QA")
        st.write(res["comment"])

# ====================== ЕКСПОРТ В EXCEL ======================
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
            
            comment_df = pd.DataFrame([["Коментар", res["comment"]]], columns=["Поле", "Значення"])
            comment_df.to_excel(writer, index=False, sheet_name=sheet_name, 
                              startrow=len(meta_df) + len(scores_df) + 4)

    xls.seek(0)
    st.download_button(
        label="📥 Завантажити результати у XLSX",
        data=xls,
        file_name="qa_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

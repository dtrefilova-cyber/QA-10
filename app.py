import streamlit as st
import pandas as pd
import requests
import json
import re
from io import BytesIO
from datetime import datetime
from openai import OpenAI

# ====================== API KEYS ======================
DEEPGRAM_API_KEY = st.secrets["DEEPGRAM_API_KEY"]
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]

client = OpenAI(api_key=OPENAI_API_KEY)

st.title("🎧 QA-10: Аналіз дзвінків")

check_date = st.date_input("Дата перевірки", datetime.today())

qa_managers_list = [
    "Дар'я", "Надя", "Настя", "Владимира", "Діана", "Руслана", "Олексій"
]

# ====================== ІНТЕРФЕЙС ======================
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
        st.error(f"Deepgram error: {response.status_code}")
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

    prompt = f"""
Ти — строгий QA-аналітик телефонних дзвінків для казино.

Сайти компанії: 777, Betking, Vegas.

=== ПРАВИЛА ПРЕЗЕНТАЦІЇ (дуже важливо!) ===
Презентація зараховується тільки якщо менеджер:
- пропонує слот, бонус, турнір, акцію, фріспіни, активність
- АБО каже, що надішле пропозицію на email / пошту
Просто згадка назви сайту ("казино 777") — НЕ є презентацією.

=== КРИТИЧНІ СИТУАЦІЇ ===
Якщо клієнт:
- за кермом, у лікарні, в зоні бойових дій
- грубо кинув слухавку
- матюкався або був дуже грубим
→ критерії "Спроба презентації", "Пропозиція бонусу" та "Домовленість про наступний контакт" вважаються виконаними на максимум (10 балів).

=== ДОМОВЛЕНІСТЬ ПРО НАСТУПНИЙ КОНТАКТ ===
- exact_time — менеджер назвала конкретний час (наприклад: "о 15:00", "через годину", "завтра о 14")
- day — домовленість тільки про день
- offer — просто "передзвоню"
- none — немає домовленості

Поверни **тільки** валідний JSON без будь-якого додаткового тексту:

{{
  "manager_introduced_self": true/false,
  "client_name_used": true/false,
  "presentation_detected": true/false,
  "bonus_offered": true/false,
  "bonus_conditions_count": число (0-3),
  "client_busy_or_rude": true/false,
  "client_hung_up": true/false,
  "manager_active": true/false,
  "followup_type": "none / offer / day / exact_time",
  "objection_detected": true/false
}}

Початок дзвінка:
{intro}

Середина дзвінка:
{middle}

Кінець дзвінка:
{ending}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o",          # рекомендується gpt-4o або gpt-4.1
            temperature=0,
            messages=[
                {"role": "system", "content": "Ти аналізатор дзвінків. Відповідай виключно JSON."},
                {"role": "user", "content": prompt}
            ]
        )

        text = response.choices[0].message.content
        match = re.search(r"\{.*\}", text, re.S | re.DOTALL)

        if match:
            features = json.loads(match.group())
        else:
            features = {}
    except:
        features = {}

    # Дефолтні значення
    defaults = {
        "manager_introduced_self": False,
        "client_name_used": False,
        "presentation_detected": False,
        "bonus_offered": False,
        "bonus_conditions_count": 0,
        "client_busy_or_rude": False,
        "client_hung_up": False,
        "manager_active": True,
        "followup_type": "none",
        "objection_detected": False
    }

    for k, v in defaults.items():
        features.setdefault(k, v)

    features["raw_text"] = dialogue.lower()
    return features


# ====================== SCORING ======================
def score_call(features, meta):
    scores = {}

    # 1. Привітання
    if features["manager_introduced_self"] and features["client_name_used"]:
        scores["Привітання"] = 5
    elif features["manager_introduced_self"] or features["client_name_used"]:
        scores["Привітання"] = 2.5
    else:
        scores["Привітання"] = 0

    scores["Дружелюбне питання / Мета дзвінка"] = 2.5
    scores["Спроба продовжити розмову"] = 5 if features["manager_active"] else 0

    # 2. Спроба презентації
    is_critical = features.get("client_busy_or_rude", False) or features.get("client_hung_up", False)
    
    if is_critical:
        scores["Спроба презентації"] = 5
    else:
        text = features.get("raw_text", "")
        presentation_keywords = ["слот", "бонус", "турнір", "акці", "фріспін", "активн", "надішл", "пошт", "email"]
        has_presentation = any(kw in text for kw in presentation_keywords)
        scores["Спроба презентації"] = 5 if has_presentation or features["presentation_detected"] else 0

    # 3. Домовленість про наступний контакт
    followup = features["followup_type"]
    if is_critical or followup == "exact_time":
        scores["Домовленість про наступний контакт"] = 10
    elif followup == "day":
        scores["Домовленість про наступний контакт"] = 7.5
    elif followup == "offer":
        scores["Домовленість про наступний контакт"] = 5
    else:
        scores["Домовленість про наступний контакт"] = 0

    # 4. Пропозиція бонусу
    if is_critical:
        scores["Пропозиція бонусу"] = 10
    else:
        bonus_conditions = features["bonus_conditions_count"]
        if not features["bonus_offered"]:
            scores["Пропозиція бонусу"] = 0
        elif bonus_conditions == 0:
            scores["Пропозиція бонусу"] = 5
        elif bonus_conditions == 1:
            scores["Пропозиція бонусу"] = 7.5
        else:
            scores["Пропозиція бонусу"] = 10

    # 5. Завершення
    scores["Завершення"] = 5 if (is_critical or features["manager_active"]) else 0

    # 6. Передзвон клієнту
    repeat = meta["repeat_call"]
    if repeat == "так, був протягом години":
        scores["Передзвон клієнту"] = 10
    elif repeat == "так, був протягом 3 годин":
        scores["Передзвон клієнту"] = 5
    else:  # "ні, не було"
        if is_critical or features.get("client_hung_up", False):
            scores["Передзвон клієнту"] = 10   # за твоєю правкою
        else:
            scores["Передзвон клієнту"] = 0

    # Інші критерії
    scores["Не додумувати"] = 5
    scores["Якість мовлення"] = meta["speech_score"]
    scores["Професіоналізм"] = 5 if meta["bonus_check"] == "помилково нараховано" else 10
    scores["CRM-картка"] = 5 if meta.get("manager_comment") else 0
    scores["Робота із запереченнями"] = 10 if not features["objection_detected"] else 5

    # Зливання клієнта
    scores["Зливання клієнта"] = 15 if (is_critical or features["manager_active"]) else 10

    return scores


# ====================== КОМЕНТАР ======================
def generate_comment(dialogue):
    prompt = f"""
Підсумуй дзвінок у 1-2 реченнях. Зазнач сильну сторону менеджера та одну ключову рекомендацію.
{dialogue}
"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            temperature=0.3,
            messages=[
                {"role": "system", "content": "Ти QA-аналітик."},
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content
    except:
        return "Не вдалося згенерувати коментар."


# ====================== GOOGLE SHEETS (якщо використовуєш) ======================
# Якщо потрібно підключення до Google Sheets — додай сюди свій код.
# Наразі залишив заглушку, щоб код запускався.

CRITERIA_ROWS = {
    "Привітання": 5, "Дружелюбне питання / Мета дзвінка": 6, "Спроба продовжити розмову": 7,
    "Спроба презентації": 8, "Домовленість про наступний контакт": 9, "Пропозиція бонусу": 10,
    "Завершення": 11, "Передзвон клієнту": 12, "Не додумувати": 13, "Якість мовлення": 14,
    "Професіоналізм": 15, "CRM-картка": 16, "Робота із запереченнями": 17, "Зливання клієнта": 18
}

META_ROWS = {"call_date": 1, "qa_manager": 2, "client_id": 3, "check_date": 4}


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
                "comment": comment,
                "features": features
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
            
            # Meta
            meta_df = pd.DataFrame(list(res["meta"].items()), columns=["Поле", "Значення"])
            meta_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)
            
            # Scores
            scores_df = pd.DataFrame(list(res["scores"].items()), columns=["Критерій", "Оцінка"])
            scores_df["Оцінка"] = scores_df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
            scores_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=len(meta_df) + 2)
            
            # Comment
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

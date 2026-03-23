import streamlit as st
import pandas as pd
import requests
import json
import re
from io import BytesIO
from openai import OpenAI

# =========================
# API Keys
# =========================
DEEPGRAM_API_KEY = st.secrets["DEEPGRAM_API_KEY"]
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]

client = OpenAI(api_key=OPENAI_API_KEY)

# =========================
# Session state
# =========================
if "dialogue" not in st.session_state:
    st.session_state.dialogue = None
if "analysis" not in st.session_state:
    st.session_state.analysis = None

# =========================
# UI
# =========================
st.title("🎧 AI Аналіз дзвінків для QA (v5.0 UI + новий промпт)")

qa_manager = st.text_input("QA менеджер")
retention_manager = st.text_input("Менеджер RET")
client_id = st.text_input("ID клієнта")
call_date = st.text_input("Дата дзвінка (ДД-ММ-РРРР)")
check_date = st.text_input("Дата перевірки (ДД-ММ-РРРР)")
bonus_check = st.selectbox("Бонус:", ["так", "ні", "не потрібно"])

repeat_call = st.selectbox(
    "Повторний дзвінок:",
    [
        "так, був протягом години",
        "так, був протягом 3 годин",
        "ні, не було"
    ]
)

manager_comment = st.text_area("Коментар менеджера", height=120)

audio_file = st.file_uploader("🎙️ Завантаж аудіофайл", type=["mp3", "wav", "m4a"])

if audio_file and st.button("Аналізувати дзвінок"):
    st.write("⏳ Обробка аудіо...")

    # Deepgram
    url = "https://api.deepgram.com/v1/listen"
    params = {"model": "nova-2","language": "uk","diarize": "true","utterances": "true","punctuate": "true","smart_format": "true"}
    headers = {"Authorization": f"Token {DEEPGRAM_API_KEY}"}
    response = requests.post(url, headers=headers, params=params, files={"audio": audio_file})
    result = response.json()

    clean_dialogue = []
    current_speaker = None
    current_text = ""
    for u in result["results"]["utterances"]:
        speaker = "Менеджер" if u["speaker"] == 0 else "Гравець"
        text = u["transcript"].strip()
        if speaker == current_speaker:
            current_text += " " + text
        else:
            if current_speaker is not None:
                clean_dialogue.append(f"{current_speaker}: {current_text}")
            current_speaker = speaker
            current_text = text
    if current_text:
        clean_dialogue.append(f"{current_speaker}: {current_text}")
    final_dialogue = "\n".join(clean_dialogue)
    st.session_state.dialogue = final_dialogue

    # OpenAI analysis
    st.write("⏳ Аналіз дзвінка за КЛН...")

    prompt = f"""
Ти — експерт з контролю якості дзвінків у казино. 
Оціни дзвінок менеджера за 14 критеріями КЛН.

Використовуй ТІЛЬКИ такі бали для кожного критерію:

- greeting: 0, 2.5, 5
- friendly_question: 0, 2.5, 5
- continue_conversation: 0, 2.5, 5
- presentation_attempt: 0, 5, 7.5, 10
- next_contact: 0, 5, 7.5, 10
- bonus_offer: 0, 5, 7.5, 10
- closing: 0, 2.5, 5
- callback: 0, 2.5, 5, 7.5, 10
- not_assume: 0, 5
- speech_quality: 0, 2.5, 5, 7.5, 10
- professionalism: 0, 5, 7.5, 10
- crm_card: 0, 2.5, 5, 7.5, 10
- objection_handling: 0, 5, 7.5, 10
- client_dumping: 0, 5, 7.5, 10, 15

Правила callback:
- Якщо repeat_call = "так, був протягом години" → 10 балів.
- Якщо repeat_call = "так, був протягом 3 годин" → 7.5 балів.
- Якщо repeat_call = "ні, не було":
    - Якщо у розмові була домовленість про повторний дзвінок → 0 балів.
    - Якщо домовленості не було → 10 балів.

Правила professionalism:
- Якщо bonus_check = "так" і у коментарі є бонус → максимальний бал.
- Якщо bonus_check = "ні", у коментарі немає бонусу і в розмові бонус не згадано → максимальний бал.
- Якщо bonus_check = "ні", але у коментарі є бонус і в розмові згадано бонус → 5 балів.
- Якщо bonus_check = "не потрібно" → максимальний бал.

Правила crm_card:
- Якщо зміст коментаря збігається зі змістом розмови → максимальний бал.
- Якщо інформація сильно розходиться → 2.5.
- Якщо коментар відсутній → 0.

Формат відповіді (строго тільки JSON):

{{
  "greeting": 0,
  "friendly_question": 0,
  "continue_conversation": 0,
  "presentation_attempt": 0,
  "next_contact": 0,
  "bonus_offer": 0,
  "closing": 0,
  "callback": 0,
  "not_assume": 0,
  "speech_quality": 0,
  "professionalism": 0,
  "crm_card": 0,
  "objection_handling": 0,
  "client_dumping": 0,
  "comment": "Коротке резюме дзвінка та поради для менеджера."
}}

Дані:
bonus_check = "{bonus_check}"
repeat_call = "{repeat_call}"
manager_comment = "{manager_comment}"

Транскрипція дзвінка:
{final_dialogue}
"""

    response = client.chat.completions.create(
        model="gpt-4.1",
        messages=[{"role": "system", "content": "Ти експерт з контролю якості дзвінків казино."},
                  {"role": "user", "content": prompt}],
        temperature=0
    )

    analysis_text = response.choices[0].message.content

    # Спроба розпарсити як JSON
    try:
        analysis_json = json.loads(analysis_text)
    except:
        analysis_json = {}

    # Постобробка allowed scores
    allowed_scores = {
        "greeting": [0, 2.5, 5],
        "friendly_question": [0, 2.5, 5],
        "continue_conversation": [0, 2.5, 5],
        "presentation_attempt": [0, 5, 7.5, 10],
        "next_contact": [0, 5, 7.5, 10],
        "bonus_offer": [0, 5, 7.5, 10],
        "closing": [0, 2.5, 5],
        "callback": [0, 2.5, 5, 7.5, 10],
        "not_assume": [0, 5],
        "speech_quality": [0, 2.5, 5, 7.5, 10],
        "professionalism": [0, 5, 7.5, 10],
        "crm_card": [0, 2.5, 5, 7.5, 10],
        "objection_handling": [0, 5, 7.5, 10],
        "client_dumping": [0, 5, 7.5, 10, 15],
    }

    clean_scores = {}
    for k, v in analysis_json.items():
        if k == "comment":
            continue
        try:
            val = float(v)
            allowed = allowed_scores.get(k, [])
            closest = min(allowed, key=lambda x: abs(x - val))
            clean_scores[k] = closest
        except:
            clean_scores[k] = 0

    clean_scores["comment"] = analysis_json.get("comment", "")

    st.session_state.analysis = clean_scores

# =========================
# Вивід результатів
# =========================
if st.session_state.analysis:
    analysis_json = st.session_state.analysis

    criteria_map = {
        "greeting": "Вітання",
        "friendly_question": "Дружнє питання",
        "continue_conversation": "Продовження розмови",
        "presentation_attempt": "Спроба презентації",
        "next_contact": "Наступний контакт",
        "bonus_offer": "Пропозиція бонусу",
        "closing": "Завершення",
        "callback": "Повторний контакт",
        "not_assume": "Не робити припущень",
        "speech_quality": "Якість мовлення",
        "professionalism": "Професіоналізм",
        "crm_card": "CRM-картка",
        "objection_handling": "Робота із запереченнями",
        "client_dumping": "Зливання клієнта"
    }

    scores = {criteria_map.get(k, k): v for k, v in analysis_json.items() if k != "comment"}
    df = pd.DataFrame(scores.items(), columns=["Критерій", "Оцінка"])

    df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}" if not float(x).is_integer() else str(int(x)))

    st.write("✅ Результати аналізу:")
    st.table(df)

    total_score = sum(float(v) for v in df["Оцінка"].apply(lambda x: x.replace(",", ".")).astype(float))
    st.markdown(f"**Загальний бал:** {total_score:.1f}")

    st.markdown("### Коментар")
    st.write(analysis_json.get("comment", ""))

    # Редагування
    st.markdown("### ✏️ Редагування оцінок")
    edited_scores = {}
    for k, v in scores.items():
        edited_scores[k] = st.number_input(f"{k}", value=float(v), step=0.5)

    if st.button("Оновити результати"):
        df = pd.DataFrame(edited_scores.items(), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}" if not float(x).is_integer() else str(int(x)))
        st.table(df)
        total_score = sum(float(v) for v in df["Оцінка"].apply(lambda x: x.replace(",", ".")).astype(float))
        st.markdown(f"**Загальний бал (оновлений):** {total_score:.1f}")

    # Формуємо службові дані
    meta_df = pd.DataFrame({
        "Поле": [
            "QA менеджер",
            "Менеджер RET",
            "ID клієнта",
            "Дата дзвінка",
            "Дата перевірки"
        ],
        "Значення": [
            qa_manager,
            retention_manager,
            client_id,
            call_date,
            check_date
        ]
    })

    # Експорт у XLSX
    xls = BytesIO()
    with pd.ExcelWriter(xls, engine="openpyxl") as writer:
        meta_df.to_excel(writer, index=False, sheet_name="Results", startrow=0)
        df.to_excel(writer, index=False, sheet_name="Results", startrow=len(meta_df) + 2)
    xls.seek(0)

    st.download_button(
        "📥 Завантажити результати у XLSX",
        xls,
        "analysis_results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

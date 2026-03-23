import streamlit as st
import pandas as pd
import requests
import json
from io import BytesIO
from datetime import datetime
from openai import OpenAI

# =========================
# API Keys
# =========================
DEEPGRAM_API_KEY = st.secrets["DEEPGRAM_API_KEY"]
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]

client = OpenAI(api_key=OPENAI_API_KEY)

st.title("🎧 QA-10: Аналіз 10 дзвінків")

# Загальна дата перевірки
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
for row in range(5):  # 5 рядків по 2 дзвінки = 10
    col1, col2 = st.columns(2)
    for col, idx in zip([col1, col2], [row*2+1, row*2+2]):
        with col.expander(f"📞 Дзвінок {idx}", expanded=False):
            audio_url = st.text_input("Посилання на аудіо", key=f"url_{idx}")
            qa_manager = st.selectbox("QA менеджер", qa_managers_list, key=f"qa_{idx}")
            ret_manager = st.text_input("Менеджер RET", key=f"ret_{idx}")
            client_id = st.text_input("ID клієнта", key=f"client_{idx}")
            call_date = st.text_input("Дата дзвінка (ДД-ММ-РРРР)", key=f"date_{idx}")
            bonus_check = st.selectbox("Бонус", ["правильно нараховано", "помилково нараховано", "не потрібно"], key=f"bonus_{idx}")
            repeat_call = st.selectbox("Повторний дзвінок", [
                "так, був протягом години",
                "так, був протягом 3 годин",
                "ні, не було"
            ], key=f"repeat_{idx}")
            manager_comment = st.text_area("Коментар менеджера", height=80, key=f"comment_{idx}")

            calls.append({
                "url": audio_url,
                "qa_manager": qa_manager,
                "ret_manager": ret_manager,
                "client_id": client_id,
                "call_date": call_date,
                "check_date": check_date.strftime("%d-%m-%Y"),
                "bonus_check": bonus_check,
                "repeat_call": repeat_call,
                "manager_comment": manager_comment
            })

# =========================
# Обробка дзвінків
# =========================
def transcribe_audio(audio_url):
    if not audio_url:
        return None
    url = "https://api.deepgram.com/v1/listen"
    params = {"model": "nova-2","language": "uk","diarize": "true","utterances": "true","punctuate": "true","smart_format": "true"}
    headers = {"Authorization": f"Token {DEEPGRAM_API_KEY}"}
    response = requests.post(url, headers=headers, params=params, json={"url": audio_url})
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
    return "\n".join(clean_dialogue)

def analyze_call(final_dialogue, meta):
    if not final_dialogue:
        return None

    prompt = f"""
Ти — експерт з контролю якості дзвінків у казино. 
Оціни дзвінок менеджера за 14 критеріями КЛН.

⚠️ Важливо:
- Відповідь має бути строго у форматі JSON, без пояснень і без тексту поза JSON.
- Не залишай усі оцінки нулями. Якщо критерій виконано частково — вибери найближче значення з дозволених.
- Використовуй тільки дозволені значення для кожного критерію:

greeting: 0, 2.5, 5
friendly_question: 0, 2.5, 5
continue_conversation: 0, 2.5, 5
presentation_attempt: 0, 5, 7.5, 10
next_contact: 0, 5, 7.5, 10
bonus_offer: 0, 5, 7.5, 10
closing: 0, 2.5, 5
callback: 0, 2.5, 5, 7.5, 10
not_assume: 0, 5
speech_quality: 0, 2.5, 5, 7.5, 10
professionalism: 0, 5, 7.5, 10
crm_card: 0, 2.5, 5, 7.5, 10
objection_handling: 0, 5, 7.5, 10
client_dumping: 0, 5, 7.5, 10, 15

Формат відповіді:
{{
  "greeting": <оцінка>,
  "friendly_question": <оцінка>,
  "continue_conversation": <оцінка>,
  "presentation_attempt": <оцінка>,
  "next_contact": <оцінка>,
  "bonus_offer": <оцінка>,
  "closing": <оцінка>,
  "callback": <оцінка>,
  "not_assume": <оцінка>,
  "speech_quality": <оцінка>,
  "professionalism": <оцінка>,
  "crm_card": <оцінка>,
  "objection_handling": <оцінка>,
  "client_dumping": <оцінка>,
  "comment": "Коротке резюме дзвінка та рекомендації"
}}

Дані:
bonus_check = "{meta['bonus_check']}"
repeat_call = "{meta['repeat_call']}"
manager_comment = "{meta['manager_comment']}"

Транскрипція дзвінка:
{final_dialogue}
"""

    response = client.chat.completions.create(
        model="gpt-4.1",
        messages=[
            {"role": "system", "content": "Ти експерт з контролю якості дзвінків казино."},
            {"role": "user", "content": prompt}
        ],
        temperature=0
    )
    return response.choices[0].message.content

# =========================
# Запуск аналізу
# =========================
if st.button("Запустити аналіз"):
    for i, call in enumerate(calls):
        if not call["url"]:
            continue
        st.write(f"⏳ Обробка дзвінка {i+1}...")
        transcript = transcribe_audio(call["url"])
        analysis_text = analyze_call(transcript, call)

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

        comment = analysis_json.get("comment", "")
        clean_scores["comment"] = comment

               # Вивід результатів
        with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):
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

            scores = {criteria_map.get(k, k): v for k, v in clean_scores.items() if k != "comment"}
            df = pd.DataFrame(scores.items(), columns=["Критерій", "Оцінка"])

            # Форматування чисел
            df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}" if not float(x).is_integer() else str(int(x)))

            st.write("✅ Результати аналізу:")
            st.table(df)

            total_score = sum(float(v) for v in df["Оцінка"].apply(lambda x: x.replace(",", ".")).astype(float))
            st.markdown(f"**Загальний бал:** {total_score:.1f}")

            st.markdown("### Коментар")
            st.write(comment)

            # Редагування
            st.markdown("### ✏️ Редагування оцінок")
            edited_scores = {}
            for k, v in scores.items():
                edited_scores[k] = st.number_input(f"{k}", value=float(v), step=0.5, key=f"edit_{i}_{k}")

            if st.button(f"Оновити результати дзвінка {i+1}"):
                df = pd.DataFrame(edited_scores.items(), columns=["Критерій", "Оцінка"])
                df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}" if not float(x).is_integer() else str(int(x)))
                st.table(df)
                total_score = sum(float(v) for v in df["Оцінка"].apply(lambda x: x.replace(",", ".")).astype(float))
                st.markdown(f"**Загальний бал (оновлений):** {total_score:.1f}")

    # =========================
    # Експорт у Excel
    # =========================
    xls = BytesIO()
    with pd.ExcelWriter(xls, engine="openpyxl") as writer:
        for i, call in enumerate(calls):
            sheet_name = f"Call_{i+1}"
            meta_df = pd.DataFrame(list(call.items()), columns=["Поле", "Значення"])
            meta_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)

    xls.seek(0)
    st.download_button(
        "📥 Завантажити результати у XLSX",
        xls,
        "qa10_results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

import streamlit as st
import pandas as pd
import requests
import json
from io import BytesIO
from datetime import datetime
from openai import OpenAI

DEEPGRAM_API_KEY = st.secrets["DEEPGRAM_API_KEY"]
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]

client = OpenAI(api_key=OPENAI_API_KEY)

st.title("🎧 QA-10: Аналіз дзвінків")

check_date = st.date_input("Дата перевірки", datetime.today())

qa_managers_list = [
    "Аліна Пронь","Дар'я Трефілова","Надія Татаренко","Анастасія Собакіна",
    "Владимира Балховська","Діана Батрак","Руслана Каленіченко","Шутов Олексій"
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
            call_date = st.text_input("Дата дзвінка (ДД-ММ-РРРР)", key=f"date_{idx}")
            bonus_check = st.selectbox("Бонус", ["правильно нараховано","помилково нараховано","не потрібно"], key=f"bonus_{idx}")
            repeat_call = st.selectbox("Повторний дзвінок", [
                "так, був протягом години","так, був протягом 3 годин","ні, не було"
            ], key=f"repeat_{idx}")
            manager_comment = st.text_area("Коментар менеджера", height=80, key=f"comment_{idx}")

            calls.append({
                "url": audio_url,"qa_manager": qa_manager,"ret_manager": ret_manager,
                "client_id": client_id,"call_date": call_date,"check_date": check_date.strftime("%d-%m-%Y"),
                "bonus_check": bonus_check,"repeat_call": repeat_call,"manager_comment": manager_comment
            })

def transcribe_audio(audio_url):
    if not audio_url: return None
    url = "https://api.deepgram.com/v1/listen"
    params = {"model":"nova-2","language":"uk","diarize":"true","utterances":"true","punctuate":"true","smart_format":"true"}
    headers = {"Authorization": f"Token {DEEPGRAM_API_KEY}"}
    response = requests.post(url, headers=headers, params=params, json={"url": audio_url})
    result = response.json()
    clean_dialogue, current_speaker, current_text = [], None, ""
    for u in result["results"]["utterances"]:
        speaker = "Менеджер" if u["speaker"] == 0 else "Гравець"
        text = u["transcript"].strip()
        if speaker == current_speaker: current_text += " " + text
        else:
            if current_speaker is not None: clean_dialogue.append(f"{current_speaker}: {current_text}")
            current_speaker, current_text = speaker, text
    if current_text: clean_dialogue.append(f"{current_speaker}: {current_text}")
    return "\n".join(clean_dialogue)

def analyze_call(final_dialogue, meta):
    if not final_dialogue:
        return None

    prompt = f"""Ти — експерт з контролю якості дзвінків у казино.
Оціни дзвінок менеджера за 14 критеріями КЛН.

⚠️ Важливо:
- Відповідь має бути строго у форматі JSON, без пояснень і без тексту поза JSON.
- Використовуй тільки дозволені значення, описані нижче.
- Не залишай усі оцінки нулями.

Правила оцінювання:
1. Привітання: якщо менеджер привітався коректно → 5; якщо частково виконав критерій → 2.5; якщо не привітався → 0.
2. Дружелюбне питання / Мета дзвінка: якщо є питання або озвучена мета дзвінка → 2.5; якщо відсутнє → 0.
3. Спроба продовжити розмову: якщо менеджер проявив ініціативу, щоб продовжити діалог → 5; якщо менеджер не спробував продовжити розмову і погодився перервати → 0.
4. Спроба презентації: якщо менеджер намагався презентувати → 10; якщо ні → 0.
5. Домовленість про наступний контакт: дата+час → 10; день/дата+частина дня → 7.5; тільки дата/день → 5; ≥2 спроби домовитись без відповіді → 10.
6. Пропозиція бонусу: не згадано → 0; згадано без умов → 5; згадано з умовами → 10.
7. Прощання: якщо менеджер попрощався коректно → 5; якщо не попрощався → 0.
8. Передзвон клієнту: «так, був протягом години» → 10; «так, був протягом 3 годин» → 5; «ні, не було» + домовленість → 0; «ні, не було» + домовленості немає → 10.
9. Не робити припущень: якщо менеджер не робив припущень → 5; якщо робив → 0.
10. Якість мовлення: чітке, зрозуміле мовлення → 2.5; середнє → 1.0; погане → 0.
11. Професіоналізм: правильний бонус → 10; неправильний бонус → 5; заборонені слова → 0.
12. CRM-картка: якщо дані у картці відповідають змісту розмови → 5; якщо частково відповідають → 2.5; якщо не відповідають або картка порожня → 0.
13. Робота із запереченнями: заперечень не було → 10; було, але проігноровано → 0.
14. Зливання клієнта: шукає причину завершити → 0; пасивний → 10; активно залучений → 15.

Формат відповіді:
{{
  "Привітання": <оцінка>,
  "Дружелюбне питання / Мета дзвінка": <оцінка>,
  "Спроба продовжити розмову": <оцінка>,
  "Спроба презентації": <оцінка>,
  "Домовленість про наступний контакт": <оцінка>,
  "Пропозиція бонусу": <оцінка>,
  "Прощання": <оцінка>,
  "Передзвон клієнту": <оцінка>,
  "Не робити припущень": <оцінка>,
  "Якість мовлення": <оцінка>,
  "Професіоналізм": <оцінка>,
  "CRM-картка": <оцінка>,
  "Робота із запереченнями": <оцінка>,
  "Зливання клієнта": <оцінка>,
  "Коментар": "Коротке резюме дзвінка та рекомендації"
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


if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("Запустити аналіз"):
    st.session_state["results"].clear()
    for i, call in enumerate(calls):
        if not call["url"]: continue
        st.write(f"⏳ Обробка дзвінка {i+1}...")
        transcript = transcribe_audio(call["url"])
        analysis_text = analyze_call(transcript, call)
        try: analysis_json = json.loads(analysis_text)
        except: analysis_json = {}
        scores = {k:v for k,v in analysis_json.items() if k!="comment"}
        comment = analysis_json.get("comment","")
        st.session_state["results"].append({"meta":call,"scores":scores,"comment":comment})

# Вивід результатів
criteria_map = {
    "greeting": "Вітання",
    "friendly_question": "Дружелюбне питання / Мета дзвінка",
    "continue_conversation": "Спроба продовжити розмову",
    "presentation_attempt": "Спроба презентації",
    "next_contact": "Домовленість про наступний контакт",
    "bonus_offer": "Пропозиція бонусу",
    "closing": "Прощання",
    "callback": "Передзвон клієнту",
    "not_assume": "Не робити припущень",
    "speech_quality": "Якість мовлення",
    "professionalism": "Професіоналізм",
    "crm_card": "CRM-картка",
    "objection_handling": "Робота із запереченнями",
    "client_dumping": "Зливання клієнта"
}

for i, res in enumerate(st.session_state["results"]):
    with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):
        # замінюємо ключі на українські назви
        scores = {criteria_map.get(k, k): v for k, v in res["scores"].items()}
        df = pd.DataFrame(scores.items(), columns=["Критерій","Оцінка"])
        # форматування з однією цифрою після коми
        df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
        st.table(df)

        total_score = sum(float(v) for v in res["scores"].values())
        st.markdown(f"**Загальний бал:** {total_score:.1f}")

        st.markdown("### Коментар")
        st.write(res["comment"])

# Експорт у Excel
if st.session_state["results"]:
    xls = BytesIO()
    with pd.ExcelWriter(xls, engine="openpyxl") as writer:
        for i, res in enumerate(st.session_state["results"]):
            sheet_name = f"Call_{i+1}"

            # метадані
            meta_df = pd.DataFrame(list(res["meta"].items()), columns=["Поле", "Значення"])
            meta_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)

            # оцінки (українською)
            scores = {criteria_map.get(k, k): v for k, v in res["scores"].items()}
            scores_df = pd.DataFrame(scores.items(), columns=["Критерій", "Оцінка"])
            scores_df["Оцінка"] = scores_df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
            scores_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=len(meta_df)+2)

            # коментар
            comment_df = pd.DataFrame([["Коментар", res["comment"]]], columns=["Поле", "Значення"])
            comment_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=len(meta_df)+len(scores_df)+4)

    xls.seek(0)
    st.download_button(
        "📥 Завантажити результати у XLSX",
        xls,
        "qa_results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

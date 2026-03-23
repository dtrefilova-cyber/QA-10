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

# Загальна дата перевірки (одна для всіх)
check_date = st.date_input("Дата перевірки", datetime.today())

# =========================
# Форма для 10 дзвінків (по 2 в рядку)
# =========================
calls = []
for row in range(5):  # 5 рядків по 2 дзвінки = 10
    col1, col2 = st.columns(2)

    # Дзвінок зліва
    with col1.expander(f"📞 Дзвінок {row*2+1}", expanded=False):
        audio_url = st.text_input(f"Посилання на аудіо {row*2+1}")
        qa_manager = st.text_input(f"QA менеджер {row*2+1}")
        ret_manager = st.text_input(f"Менеджер RET {row*2+1}")
        client_id = st.text_input(f"ID клієнта {row*2+1}")
        call_date = st.text_input(f"Дата дзвінка {row*2+1} (ДД-ММ-РРРР)")
        bonus_check = st.selectbox(f"Бонус {row*2+1}", ["так", "ні", "не потрібно"])
        repeat_call = st.selectbox(f"Повторний дзвінок {row*2+1}", [
            "так, був протягом години",
            "так, був протягом 3 годин",
            "ні, не було"
        ])
        manager_comment = st.text_area(f"Коментар менеджера {row*2+1}", height=80)

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

    # Дзвінок справа
    with col2.expander(f"📞 Дзвінок {row*2+2}", expanded=False):
        audio_url = st.text_input(f"Посилання на аудіо {row*2+2}")
        qa_manager = st.text_input(f"QA менеджер {row*2+2}")
        ret_manager = st.text_input(f"Менеджер RET {row*2+2}")
        client_id = st.text_input(f"ID клієнта {row*2+2}")
        call_date = st.text_input(f"Дата дзвінка {row*2+2} (ДД-ММ-РРРР)")
        bonus_check = st.selectbox(f"Бонус {row*2+2}", ["так", "ні", "не потрібно"])
        repeat_call = st.selectbox(f"Повторний дзвінок {row*2+2}", [
            "так, був протягом години",
            "так, був протягом 3 годин",
            "ні, не було"
        ])
        manager_comment = st.text_area(f"Коментар менеджера {row*2+2}", height=80)

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
Оціни дзвінок менеджера за 14 критеріями КЛН...
(Тут лишається твій повний промпт з правилами оцінювання)
Дані:
bonus_check = "{meta['bonus_check']}"
repeat_call = "{meta['repeat_call']}"
manager_comment = "{meta['manager_comment']}"

Транскрипція дзвінка:
{final_dialogue}
"""

    response = client.chat.completions.create(
        model="gpt-4.1",
        messages=[{"role": "system", "content": "Ти експерт з контролю якості дзвінків казино."},
                  {"role": "user", "content": prompt}],
        temperature=0
    )
    return response.choices[0].message.content

# =========================
# Запуск аналізу
# =========================
if st.button("Запустити аналіз"):
    results = []
    for i, call in enumerate(calls):
        if not call["url"]:
            continue
        st.write(f"⏳ Обробка дзвінка {i+1}...")
        transcript = transcribe_audio(call["url"])
        analysis = analyze_call(transcript, call)
        results.append({"meta": call, "transcript": transcript, "analysis": analysis})

        with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):
            st.write("📌 Дані:", call)
            st.write("📝 Транскрипт:", transcript)
            st.write("📊 Аналіз:", analysis)

    # =========================
    # Експорт у Excel
    # =========================
    if results:
        xls = BytesIO()
        with pd.ExcelWriter(xls, engine="openpyxl") as writer:
            for i, res in enumerate(results):
                df = pd.DataFrame({
                    "Transcript": [res["transcript"]],
                    "Analysis": [res["analysis"]]
                })
                meta_df = pd.DataFrame(list(res["meta"].items()), columns=["Поле", "Значення"])
                meta_df.to_excel(writer, index=False, sheet_name=f"Call_{i+1}", startrow=0)
                df.to_excel(writer, index=False, sheet_name=f"Call_{i+1}", startrow=len(meta_df)+2)
        xls.seek(0)

        st.download_button(
            "📥 Завантажити результати у XLSX",
            xls,
            "qa10_results.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

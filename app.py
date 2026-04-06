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
import anthropic

# ================= CONFIG =================
DEEPGRAM_API_KEY = st.secrets["DEEPGRAM_API_KEY"]
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]
ANTHROPIC_API_KEY = st.secrets.get("ANTHROPIC_API_KEY")

client = OpenAI(api_key=OPENAI_API_KEY)

claude_client = anthropic.Anthropic(
    api_key=ANTHROPIC_API_KEY
)

LOG_SHEET_ID = "1gElj3hB5CX86YsVQFG2M9DpfvMUMPq2lfuSNj-ylN94"

# ================= HEADER =================
st.markdown("""
<div class="card">
    <h2 style="margin:0;">🎧 QA-10</h2>
    <span style="color:#aaa;">Аналіз дзвінків</span>
</div>
""", unsafe_allow_html=True)

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
                key=f"bonus_{idx}"
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
def transcribe_audio(url):
    if not url:
        return None

    try:
        r = requests.post(
            "https://api.deepgram.com/v1/listen",
            headers={"Authorization": f"Token {DEEPGRAM_API_KEY}"},
            params={
                "model": "nova-3",
                "smart_format": "true",
                "punctuate": "true",
                "utterances": "true",
                "multichannel": "true",
                "diarize": "true",
                "language": "uk"
            },
            json={"url": url}
        )

        if r.status_code != 200:
            st.error(f"Deepgram error: {r.text}")
            return None

        data = r.json()
        channels = data.get("results", {}).get("channels", [])

        all_words = []

        for ch_index, ch in enumerate(channels):
            alternatives = ch.get("alternatives", [])
            if not alternatives:
                continue

            words = alternatives[0].get("words", [])

            for w in words:
                all_words.append({
                    "word": w.get("word", ""),
                    "start": w.get("start", 0),
                    "end": w.get("end", 0),
                    "speaker": f"ch_{ch_index}"
                })

        if not all_words:
            return None

        all_words.sort(key=lambda x: x["start"])

        dialogue = []
        current_speaker = all_words[0]["speaker"]
        current_phrase = []
        last_end = all_words[0]["end"]

        for w in all_words:
            speaker = w["speaker"]
            pause = w["start"] - last_end

            if speaker != current_speaker or pause > 0.5:
                if current_phrase:
                    dialogue.append(f"{current_speaker}: {' '.join(current_phrase)}")

                current_phrase = []
                current_speaker = speaker

            current_phrase.append(w["word"])
            last_end = w["end"]

        if current_phrase:
            dialogue.append(f"{current_speaker}: {' '.join(current_phrase)}")

        return "\n".join(dialogue)

    except Exception as e:
        st.error(f"Transcription exception: {str(e)}")
        return None


# ================= DICT =================
def load_replacements(sheet):
    try:
        data = sheet.get_all_records()
        return {
            row["raw"]: row["correct"]
            for row in data
            if row.get("raw") and row.get("correct")
        }
    except Exception:
        return {}


# ================= CLEAN =================
def clean_and_structure(dialogue, replacements):
    if not dialogue:
        return None

    dictionary_text = "\n".join([f"{k} → {v}" for k, v in replacements.items()])

    prompt = f"""
Ти отримуєш транскрипт дзвінка після speech-to-text.

Використовуй словник замін:
{dictionary_text}

Правила:
- виправляй помилки розпізнавання
- можна виправляти фрази, якщо вони зламані
- не змінюй сенс
- не скорочуй текст

Замінити:
ch_0 → Менеджер
ch_1 → Клієнт

Поверни тільки діалог.
"""

    try:
        res = client.chat.completions.create(
            model="gpt-5.4",
            temperature=0,
            messages=[
                {"role": "system", "content": "Виправляєш транскрипцію без зміни змісту."},
                {"role": "user", "content": prompt + "\n\n" + dialogue}
            ]
        )

        return res.choices[0].message.content.strip()

    except Exception as e:
        st.error(f"Cleaning error: {e}")
        return dialogue


def extract_segments(dialogue):
    lines = dialogue.split("\n")
    return "\n".join(lines[:5]), "\n".join(lines[5:-5]), "\n".join(lines[-5:])

# ================= GPT =================
# (без змін — залиш як було)

# ================= SCORING =================
# (без змін)

# ================= COMMENT =================
# (без змін)

# ================= RUN =================
if "results" not in st.session_state:
    st.session_state["results"] = []

col1, col2 = st.columns(2)
run_openai = col1.button("🚀 OpenAI", type="primary")
run_claude = col2.button("🧠 Claude")

if run_openai or run_claude:
    st.session_state["results"].clear()

    google_client = None
    replacements = {}

    try:
        google_client = connect_google()
        dict_sheet = google_client.open_by_key(LOG_SHEET_ID).worksheet("DICT")
        replacements = load_replacements(dict_sheet)
    except Exception as e:
        st.error(f"Google connect error: {e}")

    for i, call in enumerate(calls):
        if not call["url"]:
            continue

        with st.spinner(f"Аналіз дзвінка {i+1}..."):

            transcript = transcribe_audio(call["url"])
            if not transcript:
                st.warning("Немає транскрипції")
                continue

            clean_dialogue = clean_and_structure(transcript, replacements)

            if run_openai:
                features = extract_features_openai(clean_dialogue)
            else:
                features = extract_features_claude(clean_dialogue)

            if not features:
                st.warning("Помилка аналізу")
                continue

            scores = score_call(features, call)
            comment = generate_qa_comment(scores, features)

            if google_client:
                try:
                    log_sheet = google_client.open_by_key(LOG_SHEET_ID).sheet1
                    log_sheet.append_row([
                        call["check_date"],
                        call["client_id"],
                        call["ret_manager"],
                        call["url"],
                        transcript,
                        clean_dialogue,
                        comment,
                        sum(scores.values())
                    ])
                except Exception as e:
                    st.error(f"Google error: {e}")

            st.session_state["results"].append({
                "scores": scores,
                "comment": comment
            })

# ================= OUTPUT =================
for i, res in enumerate(st.session_state["results"]):
    with st.expander(f"📞 Дзвінок {i+1}", expanded=(i == 0)):
        df = pd.DataFrame(
            list(res["scores"].items()),
            columns=["Критерій", "Оцінка"]
        )
        df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
        st.table(df)

        total = sum(res["scores"].values())
        st.success(f"Загальний бал: {total:.1f}")

        st.markdown("### 💬 Коментар QA")
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

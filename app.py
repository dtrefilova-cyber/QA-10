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
                "model": "general",
                "smart_format": "true",
                "punctuate": "true",
                "detect_language": "true"
            },
            json={"url": url}
        )

        if r.status_code != 200:
            st.error(f"Deepgram error: {r.text}")
            return None

        data = r.json()
        results = data.get("results", {})

        channels = results.get("channels", [])
        utterances = results.get("utterances", [])

        if utterances:
            dialogue = []
            for u in utterances:
                speaker = f"ch_{u.get('speaker', 0)}"
                text = u.get("transcript", "")
                if text:
                    dialogue.append(f"{speaker}: {text}")
            if dialogue:
                return "\n".join(dialogue)

        if channels:
            texts = []
            for ch in channels:
                alt = ch.get("alternatives", [{}])[0]
                t = alt.get("transcript", "")
                if t:
                    texts.append(t)
            if texts:
                return "\n".join(texts)

        return None

    except Exception as e:
        st.error(f"Transcription exception: {str(e)}")
        return None

# ================= CLEAN =================
def extract_segments(dialogue):
    lines = dialogue.split("\n")
    return "\n".join(lines[:5]), "\n".join(lines[5:-5]), "\n".join(lines[-5:])

def is_autoresponder(dialogue: str) -> bool:
    if not dialogue:
        return False

    triggers = [
        "залиште повідомлення",
        "абонент недоступний",
        "voicemail",
        "please leave a message"
    ]

    return any(t in dialogue.lower() for t in triggers)

# ================= GPT =================
def extract_features_openai(dialogue, comment):
    intro, middle, ending = extract_segments(dialogue)
    prompt = get_full_analysis_prompt(intro, middle, ending, comment)

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

        if not match:
            return {}

        return json.loads(match.group())

    except Exception as e:
        st.error(f"GPT error: {e}")
        return {}

def extract_features_claude(dialogue, comment):
    intro, middle, ending = extract_segments(dialogue)
    prompt = get_full_analysis_prompt(intro, middle, ending, comment)

    try:
        response = claude_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1000,
            messages=[{"role": "user", "content": prompt}]
        )

        text = response.content[0].text
        match = re.search(r"\{[\s\S]*\}", text)

        if not match:
            return {}

        return json.loads(match.group())

    except Exception as e:
        st.error(f"Claude error: {e}")
        return {}

# ================= SCORING =================
def score_call(f, meta, dialogue):

    if is_autoresponder(dialogue):
        return {k:0 for k in [
            "Встановлення контакту","Спроба презентації","Домовленість про наступний контакт",
            "Пропозиція бонусу","Завершення розмови","Передзвон клієнту",
            "Не додумувати","Якість мовлення","Професіоналізм",
            "Оформлення картки","Утримання клієнта","Робота із запереченнями"
        ]}

    s = {}

    elements = sum([
        f.get("manager_name_present", False),
        f.get("manager_position_present", False),
        f.get("company_present", False),
        f.get("client_name_used", False),
        f.get("purpose_present", False)
    ])

    s["Встановлення контакту"] = 7.5 if elements >= 4 else 5 if elements == 3 else 2.5 if elements == 2 else 0
    s["Спроба презентації"] = f.get("presentation_score", 0)

    fup = f.get("followup_type", "none")
    s["Домовленість про наступний контакт"] = 5 if fup == "exact_time" else 2.5 if fup == "offer" else 0

    cond = len(set(f.get("bonus_conditions", [])))
    s["Пропозиція бонусу"] = 10 if cond >= 2 else 5 if cond == 1 else 0

    s["Завершення розмови"] = 5 if f.get("has_farewell") else 0

    if fup == "none":
        s["Передзвон клієнту"] = 15
    else:
        repeat = meta["repeat_call"]
        s["Передзвон клієнту"] = 15 if repeat == "так, був протягом години" else 10 if repeat == "так, був протягом 2 годин" else 0

    s["Не додумувати"] = 5
    s["Якість мовлення"] = meta["speech_score"]
    s["Професіоналізм"] = 5 if meta["bonus_check"] == "помилково нараховано" else 10

    match = f.get("comment_match_level", "none")
    complete = f.get("comment_complete", False)
    s["Оформлення картки"] = 0 if match == "none" else 2.5 if not complete else 5

    lvl = f.get("continuation_level", "none")
    s["Утримання клієнта"] = 20 if lvl == "strong" else 15 if lvl == "weak" else 10

    s["Робота із запереченнями"] = 10 if not f.get("objection_detected") else (10 if lvl == "strong" else 5 if lvl == "weak" else 0)

    return s

# ================= COMMENT =================
def generate_qa_comment(dialogue, scores):

    max_scores = {
        "Встановлення контакту": 7.5,
        "Спроба презентації": 5,
        "Домовленість про наступний контакт": 5,
        "Пропозиція бонусу": 10,
        "Завершення розмови": 5,
        "Передзвон клієнту": 15,
        "Не додумувати": 5,
        "Якість мовлення": 2.5,
        "Професіоналізм": 10,
        "Оформлення картки": 5,
        "Утримання клієнта": 20,
        "Робота із запереченнями": 10
    }

    prompt = f"""
Пройдися по кожному критерію.

якщо score == max → "виконано"
якщо score < max → коротко поясни

scores: {scores}
max: {max_scores}
"""

    try:
        res = client.chat.completions.create(
            model="gpt-5.4",
            temperature=0,
            messages=[{"role": "user", "content": prompt}]
        )
        return res.choices[0].message.content.strip()
    except:
        return "Помилка генерації коментаря"

# ================= RUN =================
if st.button("🚀 Запуск"):
    st.session_state["results"] = []

    for call in calls:

        transcript = transcribe_audio(call["url"])
        if not transcript:
            st.warning("Немає транскрипції")
            continue

        features = extract_features_openai(transcript, call["manager_comment"])
        scores = score_call(features, call, transcript)
        comment = generate_qa_comment(transcript, scores)

        st.session_state["results"].append({
            "scores": scores,
            "comment": comment
        })

# ================= OUTPUT =================
if "results" in st.session_state and st.session_state["results"]:
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

            for line in res["comment"].split("\n"):
                if line.strip():
                    st.write(line)

# ================= EXPORT =================
if "results" in st.session_state and st.session_state["results"]:
    xls = BytesIO()

    with pd.ExcelWriter(xls, engine="openpyxl") as writer:
        for i, res in enumerate(st.session_state["results"]):
            df = pd.DataFrame(
                res["scores"].items(),
                columns=["Критерій", "Оцінка"]
            )
            df.to_excel(writer, sheet_name=f"Call_{i+1}", index=False)

    xls.seek(0)

    st.download_button(
        label="📥 Завантажити Excel",
        data=xls,
        file_name="qa_results.xlsx"
    )

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

# ================= CONFIG =================
DEEPGRAM_API_KEY = st.secrets["DEEPGRAM_API_KEY"]
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]
LOG_SHEET_ID = st.secrets["LOG_SHEET_ID"]
client = OpenAI(api_key=OPENAI_API_KEY)

st.title("🎧 QA-10: Аналіз дзвінків")

check_date = st.date_input("Дата перевірки", datetime.today())

qa_managers_list = [
    "Дар'я", "Надя", "Настя", "Владимира", "Діана", "Руслана", "Олексій"
]

# ================= INPUT =================
calls = []
for row in range(5):
    col1, col2 = st.columns(2)
    for col, idx in zip([col1, col2], [row * 2 + 1, row * 2 + 2]):
        with col.expander(f"📞 Дзвінок {idx}", expanded=False):

            audio_url = st.text_input("Посилання на аудіо", key=f"url_{idx}")
            qa_manager = st.selectbox("QA менеджер", qa_managers_list, key=f"qa_{idx}")
            ret_manager = st.text_input("Менеджер RET", key=f"ret_{idx}")
            client_id = st.text_input("ID клієнта", key=f"client_{idx}")
            call_date = st.text_input("Дата дзвінка", key=f"date_{idx}")

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
def clean_transcript(text):
    replacements = {
        "вагас": "Vegas",
        "вегас": "Vegas",
        "відпрограма": "віп програма",
        "віпрограма": "віп програма",
        "артемаш": "Артем",
        "дмитроо": "Дмитро"
    }

    for k, v in replacements.items():
        text = text.replace(k, v)

    return text


def transcribe_audio(audio_url):
    if not audio_url:
        return None

    try:
        response = requests.post(
            "https://api.deepgram.com/v1/listen",
            headers={"Authorization": f"Token {DEEPGRAM_API_KEY}"},
            params={
                "model": "nova-3",
                "language": "uk",
                "punctuate": "true",
                "smart_format": "true",
                "keyterm": "Vegas,vip,віп,бонус,менеджер"
            },
            json={"url": audio_url}
        )

        if response.status_code != 200:
            st.error(f"Deepgram error: {response.text}")
            return None

        data = response.json()

    try:
        transcript = data["results"]["channels"][0]["alternatives"][0]["transcript"]
except:
        st.error("Не вдалося отримати транскрипцію")
        return None
        if not transcript.strip():
            st.warning("Порожня транскрипція")
            return None

        # 👉 очищаємо текст
        transcript = clean_transcript(transcript)

        return transcript

    except Exception as e:
        st.error(f"Помилка транскрипції: {e}")
        return None


def extract_segments(dialogue):
    lines = dialogue.split("\n")

    intro = "\n".join(lines[:5])
    middle = "\n".join(lines[5:-5]) if len(lines) > 10 else "\n".join(lines[5:])
    ending = "\n".join(lines[-5:]) if len(lines) > 5 else ""

    return intro, middle, ending

def extract_features(dialogue):
    intro, middle, ending = extract_segments(dialogue)
    prompt = get_full_analysis_prompt(intro, middle, ending)

    try:
        response = client.chat.completions.create(
            model="gpt-5.4",
            temperature=0,
            messages=[
                {"role": "system", "content": "Поверни тільки валідний JSON без тексту"},
                {"role": "user", "content": prompt}
            ]
        )

        text = response.choices[0].message.content.strip()

        match = re.search(r"\{[\s\S]*\}", text)
        if not match:
            st.error("JSON не знайдено в відповіді GPT")
            st.write(text)
            return {}

        features = json.loads(match.group())

    except Exception as e:
        st.error(f"GPT error: {e}")
        return {}

    # дефолти
    defaults = {
        "manager_name_present": False,
        "manager_position_present": False,
        "company_present": False,
        "client_name_used": False,
        "purpose_present": False,
        "bonus_offered": False,
        "bonus_conditions": [],
        "followup_type": "none",
        "objection_detected": False,
        "client_wants_to_end": False,
        "continuation_level": "none",
        "has_presentation": False,
        "has_farewell": False
    }

    for k, v in defaults.items():
        features.setdefault(k, v)

    return features

# ================= SCORING =================
def score_call(features, meta):
    scores = {}

    # 1. Контакт
    elements = sum([
        features["manager_name_present"],
        features["manager_position_present"],
        features["company_present"],
        features["client_name_used"],
        features["purpose_present"]
    ])

    scores["Встановлення контакту"] = (
        7.5 if elements >= 4 else
        5 if elements == 3 else
        2.5 if elements == 2 else
        0
    )

    # 2. Презентація
   scores["Спроба презентації"] = features.get("presentation_score", 0)

    # 3. Follow-up
    f = features.get("followup_type", "none")
    scores["Домовленість про наступний контакт"] = (
        5 if f == "exact_time" else
        2.5 if f == "offer" else
        0
    )

    # 4. Бонус (захист від дублювання)
    conditions = set(features.get("bonus_conditions", []))
    scores["Пропозиція бонусу"] = (
        0 if not features["bonus_offered"]
        else 10 if len(conditions) >= 2
        else 5
    )

    # 5. Завершення
    scores["Завершення розмови"] = 5 if features["has_farewell"] else 0

    # 6. Передзвон
    repeat = meta["repeat_call"]
    scores["Передзвон клієнту"] = (
        15 if repeat == "так, був протягом години"
        else 10 if repeat == "так, був протягом 2 годин"
        else 0
    )

    # 7. Не додумувати
    scores["Не додумувати"] = 5

    # 8. Мовлення
    scores["Якість мовлення"] = meta["speech_score"]

    # 9. Професіоналізм
    scores["Професіоналізм"] = (
        5 if meta["bonus_check"] == "помилково нараховано" else 10
    )

    # 10. CRM
    comment = meta["manager_comment"].strip()
    scores["Оформлення картки"] = (
        0 if not comment else
        2.5 if len(comment.split()) < 4 else
        5
    )

    # 11. Заперечення
    if not features["objection_detected"]:
        scores["Робота із запереченнями"] = 10
    else:
        lvl = features["continuation_level"]
        scores["Робота із запереченнями"] = (
            10 if lvl == "strong" else
            5 if lvl == "weak" else
            0
        )

    # 12. УТРИМАННЯ (КЛЮЧОВЕ ВИПРАВЛЕННЯ)
    lvl = features["continuation_level"]

    if features["client_wants_to_end"]:
    scores["Утримання клієнта"] = (
        20 if lvl == "strong" else
        15 if lvl == "weak" else
        10
    )
else:
    scores["Утримання клієнта"] = (
        20 if lvl == "strong" else
        15 if lvl == "weak" else
        10
    )

    return scores


# ================= COMMENT =================
def generate_qa_comment(scores, features):
    comments = []

    # 1. Встановлення контакту
    if scores["Встановлення контакту"] < 7.5:
        missing = []
        if not features["manager_name_present"]:
            missing.append("не назвав ім’я")
        if not features["manager_position_present"]:
            missing.append("не назвав посаду")
        if not features["company_present"]:
            missing.append("не назвав компанію")
        if not features["client_name_used"]:
            missing.append("не звернувся по імені")
        if not features["purpose_present"]:
            missing.append("не озвучив мету дзвінка")

        comments.append(f"Встановлення контакту — {', '.join(missing)}")

    # 2. Презентація
    if scores["Спроба презентації"] == 0:
        comments.append("Спроба презентації — відсутній опис продукту або гри")

    # 3. Follow-up
    if scores["Домовленість про наступний контакт"] < 5:
        comments.append("Домовленість про наступний контакт — не узгоджено точний час")

    # 4. Бонус
    if scores["Пропозиція бонусу"] < 10 and features["bonus_offered"]:
        comments.append("Пропозиція бонусу — озвучено лише одну умову бонусу")

    if scores["Пропозиція бонусу"] == 0:
        comments.append("Пропозиція бонусу — бонус не запропоновано")

    # 5. Завершення
    if scores["Завершення розмови"] < 5:
        comments.append("Завершення розмови — відсутнє коректне прощання")

    # 6. Не додумувати
    if scores["Не додумувати"] < 5:
        comments.append("Не додумувати — менеджер робив припущення замість уточнення")

    # 7. CRM
    if scores["Оформлення картки"] < 5:
        comments.append("Оформлення картки — коментар неповний або відсутній")

    # 8. Утримання
    if scores["Утримання клієнта"] < 20:
        comments.append("Утримання клієнта — слабка спроба утримати клієнта")

    # якщо все ідеально
    if not comments:
        return "Усі критерії виконані на максимальний бал"

    return "\n".join([f"- {c}" for c in comments])

# ================= RUN =================
if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("🚀 Запустити аналіз", type="primary"):
    st.session_state["results"].clear()

    # 👉 створюємо підключення до Google
    google_client = None
    try:
        google_client = connect_google()
    except Exception as e:
        st.error(f"Google Sheets error: {e}")

    for i, call in enumerate(calls):

        if not call["url"]:
            continue

        st.write(f"Обробка дзвінка {i+1}")

        transcript = transcribe_audio(call["url"])
        if not transcript:
            st.warning("Немає транскрипції")
            continue

        features = extract_features(transcript)
        scores = score_call(features, call)
        explanation = explain_scores(scores)
        comment = generate_qa_comment(scores, features)

        # 👉 запис у Google Sheets
        if google_client:
    try:
        sheet = google_client.open(call["ret_manager"]).sheet1

        # запис оцінок (як було)
        write_to_google_sheet(sheet, call, scores)

        # --- 👇 ДОДАЄМО СЮДИ ---
        start_row = 20

        existing_ids = sheet.col_values(1)[start_row-1:]
        next_row = start_row + len(existing_ids)

        sheet.update(f"A{next_row}", call["client_id"])
        sheet.update(f"B{next_row}", comment)

    except Exception as e:
        st.error(f"Помилка запису в Google Sheets: {e}")

        st.session_state["results"].append({
            "scores": scores,
            "explanation": explanation,
            "comment": comment
        })

# ================= OUTPUT =================
for i, res in enumerate(st.session_state["results"]):
    with st.expander(f"📊 Дзвінок {i+1}", expanded=True):

        df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
        st.table(df)

        total = sum(res["scores"].values())
        st.success(f"Загальний бал: {total:.1f}")

        st.markdown("### Коментар QA")
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
        "📥 Завантажити Excel",
        xls,
        file_name="qa_results.xlsx"
    )

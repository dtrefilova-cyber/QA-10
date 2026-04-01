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
                ["так, був протягом години", "так, був протягом 2 годин", "ні, не було"],
                key=f"repeat_{idx}"
            )

            manager_comment = st.text_area("Коментар менеджера", height=80, key=f"comment_{idx}")
            speech_score = st.selectbox("Якість мовлення", [2.5, 0], key=f"speech_{idx}")

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


# ================= TRANSCRIPTION =================
def transcribe_audio(audio_url):
    if not audio_url:
        return None

    url = "https://api.deepgram.com/v1/listen"
    headers = {"Authorization": f"Token {DEEPGRAM_API_KEY}"}

    response = requests.post(url, headers=headers, json={"url": audio_url})

    if response.status_code != 200:
        return None

    data = response.json()

    try:
        utterances = data["results"]["utterances"]
    except:
        return None

    dialogue = []
    for u in utterances:
        speaker = "Менеджер" if u["speaker"] == 0 else "Гравець"
        dialogue.append(f"{speaker}: {u['transcript']}")

    return "\n".join(dialogue)


def extract_segments(dialogue):
    lines = dialogue.split("\n")
    intro = "\n".join(lines[:5])
    middle = "\n".join(lines[5:-5]) if len(lines) > 10 else "\n".join(lines[5:])
    ending = "\n".join(lines[-5:]) if len(lines) > 5 else ""
    return intro, middle, ending


# ================= GPT =================
def extract_features(dialogue):
    intro, middle, ending = extract_segments(dialogue)
    prompt = get_full_analysis_prompt(intro, middle, ending)

    try:
        response = client.chat.completions.create(
            model="gpt-5.4",
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

    # Контакт
    elements = sum([
        features["manager_name_present"],
        features["manager_position_present"],
        features["company_present"],
        features["client_name_used"],
        features["purpose_present"]
    ])

    scores["Встановлення контакту"] = 7.5 if elements >= 4 else 5 if elements == 3 else 2.5 if elements == 2 else 0

    # Презентація
    scores["Спроба презентації"] = 5 if features["has_presentation"] else 0

    # Follow-up
    f = features["followup_type"]
    scores["Домовленість про наступний контакт"] = 5 if f == "exact_time" else 2.5 if f == "offer" else 0

    # Бонус
    if not features["bonus_offered"]:
        scores["Пропозиція бонусу"] = 0
    elif len(set(features["bonus_conditions"])) >= 2:
        scores["Пропозиція бонусу"] = 10
    else:
        scores["Пропозиція бонусу"] = 5

    # Завершення
    scores["Завершення розмови"] = 5 if features["has_farewell"] else 0

    # Передзвон
    repeat = meta["repeat_call"]
    scores["Передзвон клієнту"] = 15 if repeat == "так, був протягом години" else 10 if repeat == "так, був протягом 2 годин" else 0

    # Не додумує
    scores["Не додумувати"] = 5

    # Мовлення
    scores["Якість мовлення"] = meta["speech_score"]

    # Професіоналізм
    scores["Професіоналізм"] = 5 if meta["bonus_check"] == "помилково нараховано" else 10

    # CRM
    comment = meta["manager_comment"].strip()
    if not comment:
        scores["Оформлення картки"] = 0
    elif len(comment.split()) < 5:
        scores["Оформлення картки"] = 2.5
    else:
        scores["Оформлення картки"] = 5

    # Заперечення
    if not features["objection_detected"]:
        scores["Робота із запереченнями"] = 10
    else:
        lvl = features["continuation_level"]
        scores["Робота із запереченнями"] = 10 if lvl == "strong" else 5 if lvl == "weak" else 0

    # Утримання
    if not features["client_wants_to_end"]:
        scores["Утримання клієнта"] = 20
    else:
        lvl = features["continuation_level"]
        scores["Утримання клієнта"] = 20 if lvl == "strong" else 15 if lvl == "weak" else 10

    return scores


# ================= COMMENT =================
def generate_comment(dialogue):
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            temperature=0.3,
            messages=[{
                "role": "user",
                "content": f"Коротко підсумуй дзвінок: сильна сторона + що покращити\n{dialogue}"
            }]
        )
        return response.choices[0].message.content
    except:
        return ""


# ================= RUN =================
if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("🚀 Запустити аналіз"):
    st.session_state["results"].clear()

    google_client = None
    try:
        google_client = connect_google()
    except:
        pass

    for i, call in enumerate(calls):
        if not call["url"]:
            continue

        with st.spinner(f"Дзвінок {i+1}"):
            transcript = transcribe_audio(call["url"])
            if not transcript:
                continue

            features = extract_features(transcript)
            scores = score_call(features, call)
            comment = generate_comment(transcript)

            if google_client:
                try:
                    sheet = google_client.open(call["ret_manager"]).sheet1
                    write_to_google_sheet(sheet, call, scores)
                except:
                    pass

            st.session_state["results"].append({
                "scores": scores,
                "comment": comment
            })


# ================= OUTPUT =================
for i, res in enumerate(st.session_state["results"]):
    with st.expander(f"📊 Дзвінок {i+1}", expanded=True):

        df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
        st.table(df)

        total_score = sum(res["scores"].values())
        st.success(f"Загальний бал: {total_score:.1f}")

        st.markdown("### Коментар QA")
        st.write(res["comment"])

        st.markdown("### Пояснення оцінки")
        for crit, expl in res["explanation"].items():
            st.markdown(f"**{crit}:** {expl}")


# ================= EXPLANATION =================
def explain_scores(scores, features, meta):
    explanations = {}

    explanations["Встановлення контакту"] = (
        f"{scores['Встановлення контакту']:.1f} - "
        f"Елементів: {sum([features['manager_name_present'], features['manager_position_present'], features['company_present'], features['client_name_used'], features['purpose_present']])}"
    )

    explanations["Спроба презентації"] = (
        f"{scores['Спроба презентації']:.1f} - "
        f"{'Є презентація продукту' if features['has_presentation'] else 'Презентація відсутня'}"
    )

    explanations["Домовленість про наступний контакт"] = (
        f"{scores['Домовленість про наступний контакт']:.1f} - Тип: {features['followup_type']}"
    )

    explanations["Пропозиція бонусу"] = (
        f"{scores['Пропозиція бонусу']:.1f} - "
        f"Умов: {len(set(features['bonus_conditions']))}"
    )

    explanations["Завершення розмови"] = (
        f"{scores['Завершення розмови']:.1f} - "
        f"{'Було прощання' if features['has_farewell'] else 'Без прощання'}"
    )

    explanations["Передзвон клієнту"] = (
        f"{scores['Передзвон клієнту']:.1f} - {meta['repeat_call']}"
    )

    explanations["Не додумувати"] = (
        f"{scores['Не додумувати']:.1f} - Без припущень"
    )

    explanations["Якість мовлення"] = (
        f"{scores['Якість мовлення']:.1f} - Ручна оцінка"
    )

    explanations["Професіоналізм"] = (
        f"{scores['Професіоналізм']:.1f} - {meta['bonus_check']}"
    )

    comment = meta.get("manager_comment", "").strip()
    explanations["Оформлення картки"] = (
        f"{scores['Оформлення картки']:.1f} - "
        f"{'Коментар відсутній' if not comment else f'Слів: {len(comment.split())}'}"
    )

    explanations["Робота із запереченнями"] = (
        f"{scores['Робота із запереченнями']:.1f} - "
        f"Рівень: {features['continuation_level']}"
    )

    explanations["Утримання клієнта"] = (
        f"{scores['Утримання клієнта']:.1f} - "
        f"{'Клієнт хотів завершити' if features['client_wants_to_end'] else 'Клієнт не заперечував'}"
    )

    return explanations


# ================= EXPORT =================
if st.session_state["results"]:
    xls = BytesIO()

    with pd.ExcelWriter(xls, engine="openpyxl") as writer:
        for i, res in enumerate(st.session_state["results"]):
            sheet_name = f"Call_{i+1}"

            scores_df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
            scores_df["Оцінка"] = scores_df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
            scores_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)

            expl_df = pd.DataFrame(res["explanation"].items(), columns=["Критерій", "Пояснення"])
            expl_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=len(scores_df) + 2)

            comment_df = pd.DataFrame([["Коментар", res["comment"]]], columns=["Поле", "Значення"])
            comment_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=len(scores_df) + len(expl_df) + 4)

    xls.seek(0)

    st.download_button(
        label="📥 Завантажити результати у XLSX",
        data=xls,
        file_name="qa_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

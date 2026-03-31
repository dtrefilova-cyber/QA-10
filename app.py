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


# ====================== GPT ======================
def extract_features(dialogue):
    intro, middle, ending = extract_segments(dialogue)
    prompt = get_full_analysis_prompt(intro, middle, ending)
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
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
        "manager_introduced_self": False,
        "client_name_used": False,
        "bonus_offered": False,
        "bonus_conditions_count": 0,
        "followup_type": "none",
        "objection_detected": False,
        "conversation_continuation_score": 0,
        "presentation_score": 0,
        "closing_score": 0,
        "no_assumption_score": 0,
        "speech_score": 0,
        "professionalism_score": 0,
        "crm_score": 0
    }
    for k, v in defaults.items():
        features.setdefault(k, v)
    
    features["raw_text"] = dialogue.lower()
    return features


# ====================== SCORING ======================
def score_call(features, meta):
    scores = {}
    raw = features.get("raw_text", "")

    # 1. ВСТАНОВЛЕННЯ КОНТАКТУ
    has_name = features.get("manager_introduced_self", False)
    has_client = features.get("client_name_used", False)
    has_company = any(w in raw for w in ["компанія", "казино", "служба підтримки", "сайт", "проєкт"])
    has_position = any(w in raw for w in ["менеджер", "оператор", "спеціаліст"])
    has_purpose = any(w in raw[:500] for w in ["телефоную", "дзвоню", "звертаюсь", "мета", "ціль"])
    has_friendly = any(w in raw[:500] for w in ["як справ", "зручно говорити", "добрий день", "вітаю", "здрастуйте"])

    greeting_or_purpose = 1 if (has_purpose or has_friendly) else 0
    elements = sum([has_name, has_client, has_company, has_position, greeting_or_purpose])

    if elements >= 4:
        scores["Встановлення контакту"] = 7.5
    elif elements == 3:
        scores["Встановлення контакту"] = 5.0
    elif elements == 2:
        scores["Встановлення контакту"] = 2.5
    else:
        scores["Встановлення контакту"] = 0.0

    # 2. СПРОБА ПРЕЗЕНТАЦІЇ (без бонусу!)
    presentation_keywords = ["слот", "гра", "автомат", "акція", "промо", "турнір", "активність", "спін", "фріспін"]
    has_presentation = any(kw in raw for kw in presentation_keywords)

    scores["Спроба презентації"] = 5.0 if has_presentation else 0.0

    # 3. ДОМОВЛЕНІСТЬ ПРО НАСТУПНИЙ КОНТАКТ
    f = features.get("followup_type", "none")
    scores["Домовленість про наступний контакт"] = 5 if f == "exact_time" else 2.5 if f == "offer" else 0

    # 4. ПРОПОЗИЦІЯ БОНУСУ
    offered = features.get("bonus_offered", False)
    conditions = features.get("bonus_conditions_count", 0)
    if not offered:
        scores["Пропозиція бонусу"] = 0
    elif conditions >= 2:
        scores["Пропозиція бонусу"] = 10
    elif conditions >= 1:
        scores["Пропозиція бонусу"] = 5
    else:
        scores["Пропозиція бонусу"] = 0

    # 5. ЗАВЕРШЕННЯ РОЗМОВИ (0 / 2.5 / 5)
    closing = features.get("closing_score", 0)
    if closing >= 5:
        scores["Завершення розмови"] = 5.0
    elif closing >= 2.5:
        scores["Завершення розмови"] = 2.5
    else:
        scores["Завершення розмови"] = 0.0

    # 6. ПЕРЕДЗВОН КЛІЄНТУ
    repeat = meta.get("repeat_call", "")
    if repeat == "так, був протягом години":
        scores["Передзвон клієнту"] = 15
    elif repeat == "так, був протягом 2 годин":
        scores["Передзвон клієнту"] = 10
    else:
        scores["Передзвон клієнту"] = 0

    # 7. НЕ ДОДУМУЄ (0 / 2.5 / 5)
    push_phrases = ["чи є час", "чи є хвилин", "чи маєте хвилинку"]
    if any(p in raw for p in push_phrases):
        scores["Не додумувати"] = 0
    else:
        na_score = features.get("no_assumption_score", 0)
        if na_score >= 5:
            scores["Не додумувати"] = 5
        elif na_score >= 2.5:
            scores["Не додумувати"] = 2.5
        else:
            scores["Не додумувати"] = 0

    # 8. ЯКІСТЬ МОВЛЕННЯ
    scores["Якість мовлення"] = meta.get("speech_score", 0)

    # 9. ПРОФЕСІОНАЛІЗМ (0 / 5 / 10)
    forbidden_words = ["лотерея","акція","розіграш","реклама","подарунок","популяризація","лотерейний білет","даруємо","розігруємо","конкурс","кешбек","відшкодуємо","компенсація","повернення","фріспіни","безкоштовно","страхування","страховка","ставка без ризику","фрібет","бездеп"]
    if any(w in raw for w in forbidden_words):
        scores["Професіоналізм"] = 0
    elif meta.get("bonus_check") == "помилково нараховано":
        scores["Професіоналізм"] = 5
    else:
        scores["Професіоналізм"] = 10

    # 10. ОФОРМЛЕННЯ КАРТКИ В CRM (0 / 2.5 / 5)
    comment = meta.get("manager_comment", "").strip().lower()
    if not comment:
        scores["Оформлення картки"] = 0
    elif "бонус" in comment and "час" in comment:
        scores["Оформлення картки"] = 5
    elif any(kw in comment for kw in ["бонус","повтор","дзвінок","час"]):
        scores["Оформлення картки"] = 2.5
    else:
        scores["Оформлення картки"] = 0

    # 11. РОБОТА ІЗ ЗАПЕРЕЧЕННЯМИ
    if not features.get("objection_detected", False):
        scores["Робота із запереченнями"] = 10
    else:
        cont = features.get("conversation_continuation_score", 0)
        if cont == 5:
            scores["Робота із запереченнями"] = 10
        elif cont == 2.5:
            scores["Робота із запереченнями"] = 5
        else:
            scores["Робота із запереченнями"] = 0

    # 12. УТРИМАННЯ КЛІЄНТА (0 / 10 / 15 / 20)
    cont = features.get("conversation_continuation_score", 0)
    if cont == 5:
        if any(p in raw for p in ["знайдете хвилинку","можливо зараз","ще кілька хвилин"]):
            scores["Утримання клієнта"] = 20
        else:
            scores["Утримання клієнта"] = 15
    elif cont == 2.5:
        scores["Утримання клієнта"] = 15
    elif cont == 0:
        scores["Утримання клієнта"] = 0
    else:
        scores["Утримання клієнта"] = 10

    return scores


# ====================== COMMENT ======================
def generate_comment(dialogue):
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            temperature=0.3,
            messages=[{
                "role": "user",
                "content": f"Підсумуй дзвінок у 1-2 реченнях. Вкажи сильну сторону менеджера і одну рекомендацію.\n{dialogue}"
            }]
        )
        return response.choices[0].message.content
    except:
        return "Не вдалося згенерувати коментар."


def explain_scores(dialogue, scores):
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            temperature=0,
            messages=[{
                "role": "user",
                "content": f"""
Поясни коротко причину оцінки по кожному критерію.
Оцінки:
{scores}
Дзвінок:
{dialogue}
Формат:
Критерій: причина
"""
            }]
        )
        return response.choices[0].message.content
    except:
        return ""


# ====================== RUN ======================
if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("🚀 Запустити аналіз", type="primary"):
    st.session_state["results"].clear()

    try:
        google_client = connect_google()
    except Exception as e:
        st.warning(f"Не вдалося підключитись до Google Sheets: {e}")
        google_client = None

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
            explanation = explain_scores(transcript, scores)

            if google_client:
                try:
                    spreadsheet = google_client.open(call["ret_manager"])
                    sheet = spreadsheet.sheet1
                    write_to_google_sheet(sheet, call, scores)
                except Exception as e:
                    st.warning(f"Помилка запису в Google Sheets: {e}")

            st.session_state["results"].append({
                "meta": call,
                "scores": scores,
                "comment": comment,
                "explanation": explanation
            })


# ====================== OUTPUT ======================
for i, res in enumerate(st.session_state["results"]):
    with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):
        df = pd.DataFrame(list(res["scores"].items()), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(lambda x: f"{float(x):.1f}")
        st.table(df)

        total_score = sum(res["scores"].values())
        st.success(f"Загальний бал: {total_score:.1f}")

        st.markdown("### Коментар QA")
        st.write(res["comment"])
        st.markdown("### Пояснення оцінки")
        st.write(res["explanation"])


# ====================== EXPORT ======================
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

            comment_df = pd.DataFrame(
                [["Коментар", res["comment"]]],
                columns=["Поле", "Значення"]
            )
            comment_df.to_excel(
                writer,
                index=False,
                sheet_name=sheet_name,
                startrow=len(meta_df) + len(scores_df) + 4
            )

    xls.seek(0)
    st.download_button(
        label="📥 Завантажити результати у XLSX",
        data=xls,
        file_name="qa_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

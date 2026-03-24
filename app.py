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


# ---------------------------------
# СЛОВНИК КРИТЕРІЇВ
# ---------------------------------

criteria_rules = {
    "Привітання": {
        0: "Не назвав ім’я, посаду, казино і не звернувся до клієнта по імені",
        2.5: "Не назвав частину інформації або не звернувся до клієнта по імені",
        5: "Представився, назвав казино і звернувся до клієнта по імені"
    },

    "Дружелюбне питання / Мета дзвінка": {
        0: "Немає дружнього питання або мети дзвінка",
        2.5: "Є дружнє питання або озвучена мета дзвінка"
    },

    "Спроба продовжити розмову": {
        0: "Не намагався продовжити розмову",
        2.5: "Часткова спроба",
        5: "Активна спроба продовжити розмову"
    },

    "Спроба презентації": {
        0: "Не презентував слот або інфопривід",
        2.5: "Згадав слот або інфопривід без пояснення",
        5: "Назвав слот або інфопривід і коротко пояснив"
    },

    "Домовленість про наступний контакт": {
        0: "Домовленості немає",
        5: "Запропонував контакт без часу",
        7.5: "Домовленість про день",
        10: "Домовленість про точний час"
    },

    "Пропозиція бонусу": {
        0: "Бонус не запропоновано",
        5: "Бонус без умов",
        7.5: "Озвучені не всі умови",
        10: "Озвучені всі умови"
    },

    "Завершення": {
        0: "Не попрощався",
        2.5: "Попрощався"
    },

    "Передзвон клієнту": {
        0: "Не передзвонив попри домовленість",
        5: "Передзвонив до 3 годин",
        10: "Передзвонив до години або домовленості не було"
    },

    "Не додумувати": {
        0: "Менеджер робив припущення",
        2.5: "Запитав чи зручно говорити",
        5: "Не робив припущень"
    },

    "Якість мовлення": {
        0: "Багато русизмів або слів паразитів",
        2.5: "Мова чиста"
    },

    "Професіоналізм": {
        0: "Заборонені слова",
        5: "Помилка в бонусі або неактуальна інформація",
        10: "Все коректно"
    },

    "CRM-картка": {
        0: "Коментар відсутній",
        2.5: "Коментар неповний",
        5: "Коментар відповідає розмові"
    },

    "Робота із запереченнями": {
        0: "Ігнорування заперечення",
        2.5: "Шаблон без аналізу",
        5: "Шаблон з питанням",
        7.5: "Приклади без питання",
        10: "Повне опрацювання"
    },

    "Зливання клієнта": {
        0: "Шукає причину завершити",
        10: "Пасивний",
        15: "Активний у розмові"
    }
}

rules_json = json.dumps(criteria_rules, ensure_ascii=False, indent=2)


# ---------------------------------
# ВВЕДЕННЯ ДЗВІНКІВ
# ---------------------------------

calls = []

for row in range(5):

    col1, col2 = st.columns(2)

    for col, idx in zip([col1, col2], [row*2+1, row*2+2]):

        with col.expander(f"📞 Дзвінок {idx}", expanded=False):

            audio_url = st.text_input("Посилання на аудіо", key=f"url_{idx}")

            qa_manager = st.selectbox(
                "QA менеджер",
                qa_managers_list,
                key=f"qa_{idx}"
            )

            ret_manager = st.text_input("Менеджер RET", key=f"ret_{idx}")

            client_id = st.text_input("ID клієнта", key=f"client_{idx}")

            call_date = st.text_input(
                "Дата дзвінка (ДД-ММ-РРРР)",
                key=f"date_{idx}"
            )

            bonus_check = st.selectbox(
                "Бонус",
                ["правильно нараховано","помилково нараховано","не потрібно"],
                key=f"bonus_{idx}"
            )

            repeat_call = st.selectbox(
                "Повторний дзвінок",
                [
                    "так, був протягом години",
                    "так, був протягом 3 годин",
                    "ні, не було"
                ],
                key=f"repeat_{idx}"
            )

            manager_comment = st.text_area(
                "Коментар менеджера",
                height=80,
                key=f"comment_{idx}"
            )

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


# ---------------------------------
# ТРАНСКРИПЦІЯ
# ---------------------------------

def transcribe_audio(audio_url):

    if not audio_url:
        return None

    url = "https://api.deepgram.com/v1/listen"

    params = {
        "model":"nova-2",
        "language":"uk",
        "diarize":"true",
        "utterances":"true",
        "punctuate":"true",
        "smart_format":"true"
    }

    headers = {
        "Authorization": f"Token {DEEPGRAM_API_KEY}"
    }

    response = requests.post(
        url,
        headers=headers,
        params=params,
        json={"url": audio_url}
    )

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


# ---------------------------------
# GPT АНАЛІЗ
# ---------------------------------

def analyze_call(final_dialogue, meta):

    if not final_dialogue:
        return None

    prompt = f"""
Ти QA-аналітик казино.

Оціни дзвінок за 14 критеріями.

Використовуй тільки дозволені оцінки з правил.

Правила:
{rules_json}

Відповідь поверни СТРОГО у JSON.

Формат відповіді:

{{
"Привітання": 0,
"Дружелюбне питання / Мета дзвінка": 0,
"Спроба продовжити розмову": 0,
"Спроба презентації": 0,
"Домовленість про наступний контакт": 0,
"Пропозиція бонусу": 0,
"Завершення": 0,
"Передзвон клієнту": 0,
"Не додумувати": 0,
"Якість мовлення": 0,
"Професіоналізм": 0,
"CRM-картка": 0,
"Робота із запереченнями": 0,
"Зливання клієнта": 0,
"Коментар": ""
}}

Дані CRM:

bonus_check = "{meta['bonus_check']}"
repeat_call = "{meta['repeat_call']}"
manager_comment = "{meta['manager_comment']}"

Транскрипція дзвінка:

{final_dialogue}
"""

    response = client.chat.completions.create(
        model="gpt-4.1",
        messages=[
            {
                "role": "system",
                "content": "Ти система контролю якості дзвінків. Відповідаєш тільки JSON."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        temperature=0
    )

    return response.choices[0].message.content


# ---------------------------------
# ЗБЕРЕЖЕННЯ РЕЗУЛЬТАТІВ
# ---------------------------------

if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("Запустити аналіз"):

    st.session_state["results"].clear()

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

        scores = {k: v for k, v in analysis_json.items() if k != "Коментар"}

        comment = analysis_json.get("Коментар", "")

        st.session_state["results"].append({
            "meta": call,
            "scores": scores,
            "comment": comment
        })


# ---------------------------------
# ВИВІД РЕЗУЛЬТАТІВ
# ---------------------------------

def format_score(x):

    try:
        return f"{float(x):.1f}"
    except:
        return x


for i, res in enumerate(st.session_state["results"]):

    with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):

        df = pd.DataFrame(
            res["scores"].items(),
            columns=["Критерій", "Оцінка"]
        )

        df["Оцінка"] = df["Оцінка"].apply(format_score)

        st.table(df)

        total_score = sum(
            float(v)
            for v in res["scores"].values()
            if str(v).replace('.', '', 1).isdigit()
        )

        st.markdown(f"**Загальний бал:** {total_score:.1f}")

        st.markdown("### Коментар")

        st.write(res["comment"])


# ---------------------------------
# ЕКСПОРТ У EXCEL
# ---------------------------------

if st.session_state["results"]:

    xls = BytesIO()

    with pd.ExcelWriter(xls, engine="openpyxl") as writer:

        for i, res in enumerate(st.session_state["results"]):

            sheet_name = f"Call_{i+1}"

            meta_df = pd.DataFrame(
                list(res["meta"].items()),
                columns=["Поле","Значення"]
            )

            meta_df.to_excel(
                writer,
                index=False,
                sheet_name=sheet_name,
                startrow=0
            )

            scores_df = pd.DataFrame(
                res["scores"].items(),
                columns=["Критерій","Оцінка"]
            )

            scores_df["Оцінка"] = scores_df["Оцінка"].apply(format_score)

            scores_df.to_excel(
                writer,
                index=False,
                sheet_name=sheet_name,
                startrow=len(meta_df)+2
            )

            comment_df = pd.DataFrame(
                [["Коментар", res["comment"]]],
                columns=["Поле","Значення"]
            )

            comment_df.to_excel(
                writer,
                index=False,
                sheet_name=sheet_name,
                startrow=len(meta_df)+len(scores_df)+4
            )

    xls.seek(0)

    st.download_button(
        "📥 Завантажити результати у XLSX",
        xls,
        "qa_results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

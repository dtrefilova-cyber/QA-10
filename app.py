import streamlit as st
import pandas as pd
import requests
import json
from io import BytesIO
from datetime import datetime
from openai import OpenAI

# API keys
DEEPGRAM_API_KEY = st.secrets["DEEPGRAM_API_KEY"]
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]

client = OpenAI(api_key=OPENAI_API_KEY)

st.title("🎧 QA-10: Аналіз дзвінків")

check_date = st.date_input("Дата перевірки", datetime.today())

qa_managers_list = [
    "Аліна Пронь","Дар'я Трефілова","Надія Татаренко","Анастасія Собакіна",
    "Владимира Балховська","Діана Батрак","Руслана Каленіченко","Шутов Олексій"
]

# ------------------------
# 1. Словник правил
# ------------------------
criteria_rules = {
    "Привітання": {
        "0": "Менеджер не назвав ім’я, посаду, назву казино та не звернувся до клієнта на ім’я",
        "2.5": "Менеджер назвав лише частину (ім’я/посаду/казино) або не звернувся на ім’я",
        "5": "Менеджер назвав ім’я, посаду, назву казино та звернувся до клієнта на ім’я — завжди 5"
    },
    "Дружелюбне питання / Мета дзвінка": {
        "0": "Відсутнє дружнє питання і не озвучена мета дзвінка",
        "2.5": "Менеджер задав дружнє питання або озвучив мету дзвінка (наприклад, дзвонить щоб познайомитися чи лишити бонус)"
    },
    "Спроба продовжити розмову": {
        "0": "Менеджер не спробував продовжити розмову",
        "2.5": "Є часткова спроба, але не доведена до кінця",
        "5": "Менеджер успішно продовжив розмову"
    },
    "Спроба презентації": {
        "0": "Менеджер не презентував жодного інфоприводу чи слота з сайту",
        "2.5": "Менеджер згадав інфопривід чи слот, але без пояснення",
        "5": "Менеджер презентував інфопривід чи слот з поясненням",
        "⚠️": "Бонус ніколи не рахується як презентація. Для бонусу є окремий критерій."
    },
    "Домовленість про наступний контакт": {
        "0": "Менеджер не домовився про повторну комунікацію",
        "5": "Менеджер домовився про повторну комунікацію, але без конкретного часу",
        "7.5": "Менеджер домовився про день/дату, але не точний час",
        "10": "Менеджер домовився про конкретний час"
    },
    "Пропозиція бонусу": {
        "0": "Бонус не запропоновано",
        "5": "Запропоновано без умов",
        "7.5": "Запропоновано з неповними умовами",
        "10": "Запропоновано з усіма умовами (термін дії, мінімальний депозит, вейджер)"
    },
    "Завершення": {
        "0": "Менеджер не попрощався",
        "2.5": "Менеджер попрощався"
    },
    "Передзвон клієнту": {
        "0": "Не передзвонив, хоча була домовленість",
        "5": "Дзвінок протягом 3 годин",
        "10": "Дзвінок протягом години або домовленості не було"
    },
    "Не додумувати": {
        "0": "Менеджер робив припущення",
        "2.5": "Запитав чи зручно говорити (негатив)",
        "5": "Не додумував нічого"
    },
    "Якість мовлення": {
        "0": "Багато русизмів, паразитів",
        "2.5": "Мова чиста або незначні паразити"
    },
    "Професіоналізм": {
        "0": "Використав заборонені слова",
        "5": "Помилка у бонусі або неактуальна інформація",
        "10": "Все коректно, без помилок",
        "заборонені_слова": [
            "лотерея","акція","реклама","розіграш","даруємо","подарунок",
            "популяризація","лотерейний білет","розігруємо","конкурс","кешбек",
            "відшкодуємо","фріспіни","безкоштовно","страхування","страховка",
            "ставка без ризику","фрібет","бонуси","бонусна програма","бездеп"
        ]
    },
    "CRM-картка": {
        "0": "Коментар відсутній",
        "2.5": "Коментар неповний або не співпадає",
        "5": "Коментар повний і співпадає"
    },
    "Робота із запереченнями": {
        "0": "Ігнорування заперечення",
        "2.5": "Шаблонне опрацювання без питання",
        "5": "Шаблонне опрацювання з питанням або одне заперечення проігноровано",
        "7.5": "Опрацювання з прикладами, але без уточнюючого питання",
        "10": "Опрацював і поставив питання або заперечення не було"
    },
    "Зливання клієнта": {
        "0": "Менеджер шукає причину завершити",
        "10": "Менеджер пасивний",
        "15": "Менеджер активно залучений"
    }
}


# ------------------------
# 2. Допоміжні функції
# ------------------------
def build_prompt(criteria_rules):
    prompt_parts = []
    prompt_parts.append("Ти — експерт з контролю якості дзвінків у казино.")
    prompt_parts.append("Оціни дзвінок менеджера за 14 критеріями КЛН.")
    prompt_parts.append("⚠️ Важливо: Відповідь має бути строго у форматі JSON, без пояснень і без тексту поза JSON.")
    prompt_parts.append("У відповіді мають бути всі 14 критеріїв + 'Коментар'. Якщо критерій не застосовується — все одно постав оцінку (наприклад, 0).")

    # додаємо правила
    for criterion, rules in criteria_rules.items():
        prompt_parts.append(f"\n{criterion}:")
        for score, description in rules.items():
            if isinstance(description, list):
                prompt_parts.append(f"  {score}: {', '.join(description)}")
            else:
                prompt_parts.append(f"  {score} - {description}")

    # блок типових помилок
    prompt_parts.append("""
⚠️ Типові помилки, яких треба уникати:
1. Якщо менеджер назвав ім’я, посаду, проєкт і звернувся на ім’я — завжди 5 балів за 'Привітання'.
2. Якщо менеджер озвучив мету дзвінка — завжди 2.5 за 'Дружелюбне питання / Мета дзвінка'.
3. Часткова спроба продовжити розмову = 2.5, повна = 5.
4. Бонус НІКОЛИ не є презентацією. У 'Спроба презентації' враховуються тільки слоти чи активності з сайту.
5. Якщо домовленість про наступний контакт є, але без точного часу — завжди 5.
6. Якщо озвучені всі умови бонусу (термін дії, мінімальний депозит, вейджер) — завжди 10.
""")

    # приклад відповіді
    prompt_parts.append("""
Приклад відповіді:
{
  "Привітання": 5,
  "Дружелюбне питання / Мета дзвінка": 2.5,
  "Спроба продовжити розмову": 2.5,
  "Спроба презентації": 0,
  "Домовленість про наступний контакт": 5,
  "Пропозиція бонусу": 10,
  "Завершення": 2.5,
  "Передзвон клієнту": 0,
  "Не додумувати": 5,
  "Якість мовлення": 2.5,
  "Професіоналізм": 10,
  "CRM-картка": 5,
  "Робота із запереченнями": 7.5,
  "Зливання клієнта": 15,
  "Коментар": "Менеджер привітався, озвучив мету дзвінка, запропонував бонус з усіма умовами, домовився про повторний контакт."
}
""")

    return "\n".join(prompt_parts)


def validate_scores(analysis_json, criteria_rules):
    validated = {}
    for criterion in criteria_rules.keys():
        if criterion in analysis_json:
            validated[criterion] = analysis_json[criterion]
        else:
            validated[criterion] = 0
    validated["Коментар"] = analysis_json.get("Коментар", "")
    return validated


def transcribe_audio(audio_url):
    if not audio_url:
        return None
    url = "https://api.deepgram.com/v1/listen"
    params = {"model":"nova-2","language":"uk","diarize":"true","utterances":"true","punctuate":"true","smart_format":"true"}
    headers = {"Authorization": f"Token {DEEPGRAM_API_KEY}"}
    response = requests.post(url, headers=headers, params=params, json={"url": audio_url})
    result = response.json()
    clean_dialogue, current_speaker, current_text = [], None, ""
    for u in result["results"]["utterances"]:
        speaker = "Менеджер" if u["speaker"] == 0 else "Гравець"
        text = u["transcript"].strip()
        if speaker == current_speaker:
            current_text += " " + text
        else:
            if current_speaker is not None:
                clean_dialogue.append(f"{current_speaker}: {current_text}")
            current_speaker, current_text = speaker, text
    if current_text:
        clean_dialogue.append(f"{current_speaker}: {current_text}")
    return "\n".join(clean_dialogue)


def analyze_call(final_dialogue, meta, criteria_rules):
    if not final_dialogue:
        return None
    prompt = build_prompt(criteria_rules)
    prompt += f"""
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
    try:
        analysis_json = json.loads(response.choices[0].message.content)
    except:
        analysis_json = {}
    return validate_scores(analysis_json, criteria_rules)


def format_score(x):
    try:
        return f"{float(x):.1f}"
    except:
        return x

# ------------------------
# 3. Інтерфейс Streamlit
# ------------------------
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
            repeat_call = st.selectbox("Повторний дзвінок", ["так, був протягом години","так, був протягом 3 годин","ні, не було"], key=f"repeat_{idx}")
            manager_comment = st.text_area("Коментар менеджера", height=80, key=f"comment_{idx}")

            calls.append({
                "url": audio_url,"qa_manager": qa_manager,"ret_manager": ret_manager,
                "client_id": client_id,"call_date": call_date,"check_date": check_date.strftime("%d-%m-%Y"),
                "bonus_check": bonus_check,"repeat_call": repeat_call,"manager_comment": manager_comment
            })

if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("Запустити аналіз"):
    st.session_state["results"].clear()
    for i, call in enumerate(calls):
        if not call["url"]:
            continue
        st.write(f"⏳ Обробка дзвінка {i+1}...")
        transcript = transcribe_audio(call["url"])
        analysis = analyze_call(transcript, call, criteria_rules)
        st.session_state["results"].append({
            "meta": call,
            "scores": {k: v for k, v in analysis.items() if k != "Коментар"},
            "comment": analysis["Коментар"]
        })

# Вивід результатів
for i, res in enumerate(st.session_state["results"]):
    with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):
        df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(format_score)
        st.table(df)

        total_score = sum(
            float(v) for v in res["scores"].values() if str(v).replace('.', '', 1).isdigit()
        )
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

            # оцінки
            scores_df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
            scores_df["Оцінка"] = scores_df["Оцінка"].apply(format_score)
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


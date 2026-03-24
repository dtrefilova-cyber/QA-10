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

# -----------------------------
# Словник критеріїв
# -----------------------------
criteria_rules = {
    "Привітання": {"0": "Не назвав", "2.5": "Частково", "5": "Повністю"},
    "Дружелюбне питання / Мета дзвінка": {"0": "Відсутнє", "2.5": "Є"},
    "Спроба продовжити розмову": {"0": "Немає", "2.5": "Частково", "5": "Є"},
    "Спроба презентації": {"0": "Немає", "2.5": "Згадав", "5": "Презентував", "⚠️": "Бонус ≠ презентація"},
    "Домовленість про наступний контакт": {"0": "Немає", "5": "Є без часу", "7.5": "Є дата", "10": "Є час"},
    "Пропозиція бонусу": {"0": "Немає", "5": "Без умов", "7.5": "Неповні умови", "10": "Повні умови"},
    "Завершення": {"0": "Не попрощався", "2.5": "Попрощався"},
    "Передзвон клієнту": {"0": "Не передзвонив", "5": "До 3 годин", "10": "До години або домовленості не було"},
    "Не додумувати": {"0": "Робив припущення", "2.5": "Запитав чи зручно", "5": "Не додумував"},
    "Якість мовлення": {"0": "Русизми", "2.5": "Чиста мова"},
    "Професіоналізм": {"0": "Заборонені слова", "5": "Помилка у бонусі", "10": "Все коректно"},
    "CRM-картка": {"0": "Немає", "2.5": "Неповна", "5": "Повна"},
    "Робота із запереченнями": {"0": "Ігнорування", "2.5": "Шаблон без питання", "5": "Шаблон з питанням", "7.5": "Приклади без питання", "10": "Опрацював і поставив питання"},
    "Зливання клієнта": {"0": "Шукає причину завершити", "10": "Пасивний", "15": "Активний"}
}

# -----------------------------
# Транскрипція
# -----------------------------
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

# -----------------------------
# Аналіз дзвінка
# -----------------------------
def analyze_call(final_dialogue, meta):
    if not final_dialogue:
        return None

    prompt = "Ти — експерт з контролю якості дзвінків у казино.\n"
    prompt += "Оціни дзвінок менеджера за 14 критеріями КЛН.\n"
    prompt += "⚠️ Важливо: Відповідь має бути строго у форматі JSON.\n"

    for criterion, rules in criteria_rules.items():
        prompt += f"\n{criterion}:\n"
        for score, description in rules.items():
            prompt += f"  {score} - {description}\n"

    prompt += """
⚠️ Типові помилки:
1. Повноцінне привітання = 5.
2. Озвучена мета дзвінка = 2.5.
3. Часткова спроба продовжити = 2.5.
4. Бонус ≠ презентація.
5. Домовленість без часу = 5.
6. Бонус з усіма умовами = 10.
"""

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
    return response.choices[0].message.content

# -----------------------------
# Збереження результатів
# -----------------------------
if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("Запустити аналіз"):
    st.session_state["results"].clear()
    for i, call in enumerate(calls):
        if not call["url"]: continue
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

# -----------------------------
# Вивід результатів
# -----------------------------
def format_score(x):
    try:
        return f"{float(x):.1f}"
    except:
        return x

for i, res in enumerate(st.session_state["results"]):
    with st.expander(f"📊 Результат дзвінка {i+1}", expanded=True):
        df = pd.DataFrame(res["scores"].items(), columns=["Критерій", "Оцінка"])
        df["Оцінка"] = df["Оцінка"].apply(format_score)
        st.table(df)

        total_score = sum(float(v) for v in res["scores"].values() if str(v).replace('.', '', 1).isdigit())
        st.markdown(f"**Загальний бал:** {total_score:.1f}")

        st.markdown("### Коментар")
        st.write(res["comment"])

# -----------------------------
# Експорт у Excel
# -----------------------------
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

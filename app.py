import streamlit as st
import requests
import json
from openai import OpenAI

# -----------------------------
# Ключі API
# -----------------------------
DEEPGRAM_API_KEY = "YOUR_DEEPGRAM_KEY"
OPENAI_API_KEY = "YOUR_OPENAI_KEY"

client = OpenAI(api_key=OPENAI_API_KEY)

# -----------------------------
# Критерії
# -----------------------------
criteria_rules = {
    "Привітання": {
        "0": "Менеджер не назвав ім’я, посаду, назву казино та не звернувся до клієнта на ім’я",
        "2.5": "Менеджер назвав лише частину (ім’я/посаду/казино) або не звернувся на ім’я",
        "5": "Менеджер назвав ім’я, посаду, назву казино та звернувся до клієнта на ім’я — завжди 5"
    },
    "Дружелюбне питання / Мета дзвінка": {
        "0": "Відсутнє дружнє питання і не озвучена мета дзвінка",
        "2.5": "Менеджер задав дружнє питання або озвучив мету дзвінка"
    },
    "Спроба продовжити розмову": {
        "0": "Менеджер не спробував продовжити розмову",
        "2.5": "Є часткова спроба, але не доведена до кінця",
        "5": "Менеджер успішно продовжив розмову"
    },
    "Спроба презентації": {
        "0": "Менеджер не презентував інфопривід чи слот",
        "2.5": "Згадав, але без пояснення",
        "5": "Назвав і пояснив інфопривід чи слот",
        "⚠️": "Бонус ніколи не рахується як презентація"
    },
    "Домовленість про наступний контакт": {
        "0": "Не домовився про повторну комунікацію",
        "5": "Домовився, але без конкретного часу",
        "7.5": "Домовився про день/дату, але не точний час",
        "10": "Домовився про конкретний час"
    },
    "Пропозиція бонусу": {
        "0": "Бонус не запропоновано",
        "5": "Запропоновано без умов",
        "7.5": "Запропоновано з неповними умовами",
        "10": "Запропоновано з усіма умовами (термін дії, мінімальний депозит, вейджер)"
    },
    "Завершення": {"0": "Менеджер не попрощався", "2.5": "Менеджер попрощався"},
    "Передзвон клієнту": {"0": "Не передзвонив", "5": "Протягом 3 годин", "10": "Протягом години або домовленості не було"},
    "Не додумувати": {"0": "Менеджер робив припущення", "2.5": "Запитав чи зручно говорити", "5": "Не додумував нічого"},
    "Якість мовлення": {"0": "Багато русизмів", "2.5": "Мова чиста"},
    "Професіоналізм": {"0": "Заборонені слова", "5": "Помилка у бонусі", "10": "Все коректно"},
    "CRM-картка": {"0": "Коментар відсутній", "2.5": "Коментар неповний", "5": "Коментар повний"},
    "Робота із запереченнями": {"0": "Ігнорування", "2.5": "Шаблон без питання", "5": "Шаблон з питанням", "7.5": "Приклади без питання", "10": "Опрацював і поставив питання"},
    "Зливання клієнта": {"0": "Шукає причину завершити", "10": "Пасивний", "15": "Активно залучений"}
}

# -----------------------------
# Побудова промпту
# -----------------------------
def build_prompt(criteria_rules):
    prompt_parts = []
    prompt_parts.append("Ти — експерт з контролю якості дзвінків у казино.")
    prompt_parts.append("Оціни дзвінок менеджера за 14 критеріями КЛН.")
    prompt_parts.append("⚠️ Важливо: Відповідь має бути строго у форматі JSON.")
    for criterion, rules in criteria_rules.items():
        prompt_parts.append(f"\n{criterion}:")
        for score, description in rules.items():
            prompt_parts.append(f"  {score} - {description}")
    prompt_parts.append("""
⚠️ Типові помилки:
1. Повноцінне привітання = 5.
2. Озвучена мета дзвінка = 2.5.
3. Часткова спроба продовжити = 2.5.
4. Бонус ≠ презентація.
5. Домовленість без часу = 5.
6. Бонус з усіма умовами = 10.
""")
    return "\n".join(prompt_parts)

# -----------------------------
# Транскрипція
# -----------------------------
def transcribe_audio(audio_url):
    if not audio_url:
        return None
    url = "https://api.deepgram.com/v1/listen"
    params = {
        "model": "general",
        "tier": "enhanced",
        "language": "uk",
        "diarize": "true",
        "utterances": "true",
        "punctuate": "true",
        "smart_format": "true"
    }
    headers = {"Authorization": f"Token {DEEPGRAM_API_KEY}"}
    response = requests.post(url, headers=headers, params=params, json={"url": audio_url})
    result = response.json()
    if "utterances" in result.get("results", {}):
        clean_dialogue = []
        current_speaker, current_text = None, ""
        for u in result["results"]["utterances"]:
            speaker = "Менеджер" if u.get("speaker") == 0 else "Гравець"
            text = u.get("transcript", "").strip()
            if speaker == current_speaker:
                current_text += " " + text
            else:
                if current_speaker is not None:
                    clean_dialogue.append(f"{current_speaker}: {current_text}")
                current_speaker, current_text = speaker, text
        if current_text:
            clean_dialogue.append(f"{current_speaker}: {current_text}")
        return "\n".join(clean_dialogue)
    try:
        return result["results"]["channels"][0]["alternatives"][0]["transcript"]
    except Exception:
        return ""

# -----------------------------
# Аналіз дзвінка
# -----------------------------
def analyze_call(transcript, call, criteria_rules):
    prompt = build_prompt(criteria_rules) + "\n\nТранскрипт:\n" + transcript
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0
    )
    raw = response.choices[0].message.content
    try:
        analysis = json.loads(raw)
    except Exception:
        st.write("⚠️ Не вдалося розпарсити JSON, відповідь була:")
        st.text(raw)
        analysis = {}
    return analysis

# -----------------------------
# Streamlit UI
# -----------------------------
st.title("QA Аналіз дзвінків")

if "results" not in st.session_state:
    st.session_state["results"] = []

audio_url = st.text_input("Встав URL аудіо дзвінка")

if st.button("Запустити аналіз"):
    st.session_state["results"].clear()
    if audio_url:
        st.write("⏳ Обробка дзвінка...")
        transcript = transcribe_audio(audio_url)
        st.markdown("### Транскрипція")
        st.text(transcript)
        analysis = analyze_call(transcript, {"url": audio_url}, criteria_rules)
        if isinstance(analysis, dict):
            st.session_state["results"].append({
                "meta": {"url": audio_url},
                "scores": {k: v for k, v in analysis.items() if k != "Коментар"},
                "comment": analysis.get("Коментар", "")
            })
        else:
            st.write("⚠️ Аналіз не повернув словник")

if st.session_state["results"]:
    st.markdown("## Результати аналізу")
    for i, result in enumerate(st.session_state["results"], start=1):
        st.markdown(f"### Дзвінок {i}")
        st.json(result["scores"])
        st.markdown(f"**Коментар:** {result['comment']}")
        else:
            st.write("⚠️ Аналіз не повернув словник")

# -----------------------------
# Streamlit UI
# -----------------------------
st.title("QA Аналіз дзвінків")

if "results" not in st.session_state:
    st.session_state["results"] = []

audio_url = st.text_input("Встав URL аудіо дзвінка")

if st.button("Запустити аналіз"):
    st.session_state["results"].clear()
    if audio_url:
        st.write("⏳ Обробка дзвінка...")
        transcript = transcribe_audio(audio_url)
        st.markdown("### Транскрипція")
        st.text(transcript)
        analysis = analyze_call(transcript, {"url": audio_url}, criteria_rules)
        if isinstance(analysis, dict):
            st.session_state["results"].append({
                "meta": {"url": audio_url},
                "scores": {k: v for k, v in analysis.items() if k != "Коментар"},
                "comment": analysis.get("Коментар", "")
            })
        else:
            st.write("⚠️ Аналіз не повернув словник")

# -----------------------------
# Вивід результатів
# -----------------------------
if st.session_state["results"]:
    st.markdown("## Результати аналізу")

    import pandas as pd
    rows = []
    for i, result in enumerate(st.session_state["results"], start=1):
        scores = result["scores"].copy()
        scores["Коментар"] = result["comment"]
        scores["Дзвінок"] = i
        rows.append(scores)

    df = pd.DataFrame(rows)

    # Показуємо таблицю
    st.dataframe(df)

    # Кнопка для завантаження у Excel
    def convert_df_to_excel(df):
        from io import BytesIO
        import xlsxwriter
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Results')
        writer.close()
        return output.getvalue()

    excel_file = convert_df_to_excel(df)
    st.download_button(
        label="📥 Завантажити результати в Excel",
        data=excel_file,
        file_name="qa_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


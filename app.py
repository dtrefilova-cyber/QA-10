import streamlit as st
import pandas as pd
import requests
import json
import re
from google_sheets import (
    append_manager_log,
    append_qa_log,
    connect_google,
    load_managers_config,
    write_to_google_sheet,
)
from io import BytesIO
from datetime import datetime
from openai import OpenAI
from prompts import get_full_analysis_prompt, get_qa_comment_prompt
from prompts import get_full_analysis_prompt_claude, get_full_analysis_prompt_openai
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
DICT_SHEET_ID = "1gElj3hB5CX86YsVQFG2M9DpfvMUMPq2lfuSNj-ylN94"
KB_SHEET_ID = "1yZbtao1P1Xa0r6ZJAnjkJWikxcWQ90XbXvaT7EWQKeU"

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

def get_managers_config():
    google_client = connect_google()
    return load_managers_config(google_client, LOG_SHEET_ID)


managers_meta = {
    "headers": [],
    "header_row_index": None,
    "raw_rows_count": 0,
    "valid_rows_count": 0
}

try:
    managers_payload = get_managers_config()
    managers_config = managers_payload.get("managers", [])
    managers_meta = {
        "headers": managers_payload.get("headers", []),
        "header_row_index": managers_payload.get("header_row_index"),
        "raw_rows_count": managers_payload.get("raw_rows_count", 0),
        "valid_rows_count": managers_payload.get("valid_rows_count", 0)
    }
except Exception as e:
    managers_config = []
    st.error(f"Помилка завантаження менеджерів: {e}")

projects_list = sorted({item["project"] for item in managers_config})

if not managers_config:
    st.warning(
        "Список проєктів і менеджерів не завантажився з аркуша MANAGERS. "
        "Перевірте, що в аркуші є заголовки MANAGERS_NAME, PROJECT, SHEET_ID "
        "і що в колонці SHEET_ID заповнені значення."
    )
    st.caption(
        f"Діагностика: headers={managers_meta['headers']}, "
        f"header_row={managers_meta['header_row_index']}, "
        f"raw_rows={managers_meta['raw_rows_count']}, "
        f"valid_rows={managers_meta['valid_rows_count']}"
    )

# ================= INPUT =================
calls = []
for row in range(5):
    col1, col2 = st.columns(2)
    for col, idx in zip([col1, col2], [row * 2 + 1, row * 2 + 2]):
        with col.expander(f"📞 Дзвінок {idx}"):
            audio_url = st.text_input("Посилання", key=f"url_{idx}")
            qa_manager = st.selectbox("QA", qa_managers_list, key=f"qa_{idx}")
            selected_project = st.selectbox(
                "Проєкт",
                projects_list,
                index=None,
                placeholder="Оберіть проєкт",
                key=f"project_{idx}",
                disabled=not projects_list
            )
            project_managers = [
                item for item in managers_config
                if item["project"] == selected_project
            ]
            manager_names = [item["manager_name"] for item in project_managers]
            selected_manager = st.selectbox(
                "Менеджер РЕТ",
                manager_names,
                index=None,
                placeholder="Оберіть менеджера",
                key=f"ret_{idx}",
                disabled=not manager_names
            )
            selected_manager_data = next(
                (item for item in project_managers if item["manager_name"] == selected_manager),
                None
            )
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

            calls.append({
                "url": audio_url.strip(),
                "qa_manager": qa_manager,
                "project": selected_project or "",
                "ret_manager": selected_manager or "",
                "ret_sheet_id": selected_manager_data["sheet_id"] if selected_manager_data else "",
                "client_id": client_id,
                "call_date": call_date,
                "check_date": check_date.strftime("%d-%m-%Y"),
                "bonus_check": bonus_check,
                "repeat_call": repeat_call,
                "manager_comment": manager_comment,
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
        results = data.get("results", {})

        channels = results.get("channels", [])
        utterances = results.get("utterances", [])

        all_words = []

        # fallback якщо нема channels, але є utterances
        if not channels and utterances:
            dialogue = []
            for u in utterances:
                speaker = f"ch_{u.get('speaker', 0)}"
                text = u.get("transcript", "")
                if text:
                    dialogue.append(f"{speaker}: {text}")
            return "\n".join(dialogue)

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
def normalize_sheet_headers(row):
    return {
        str(key).strip().upper(): value
        for key, value in row.items()
    }


def load_replacements(sheet):
    try:
        data = [normalize_sheet_headers(row) for row in sheet.get_all_records()]
        return {
            str(row["RAW"]).strip(): str(row["CORRECT"]).strip()
            for row in data
            if row.get("RAW") and row.get("CORRECT")
        }
    except Exception:
        return {}


def load_kb_data(sheet):
    try:
        return [normalize_sheet_headers(row) for row in sheet.get_all_records()]
    except Exception:
        return []

import re

def apply_replacements(text, replacements):
    if not text:
        return text

    for k, v in replacements.items():
        pattern = re.compile(rf"{re.escape(k)}", re.IGNORECASE)
        text = pattern.sub(v, text)

    return text

def detect_presentation(dialogue, kb_data):
    if not dialogue:
        return False

    text = dialogue.lower()

    for row in kb_data:
        name = (row.get("NAME") or "").lower()
        aliases = (row.get("ALIASES") or "").lower().split(";")

        variants = [name] + aliases

        for v in variants:
            v = v.strip()
            if v and v in text:
                return True

    return False


def build_kb_context(kb_data):
    lines = []

    for row in kb_data:
        name = str(row.get("NAME", "")).strip()
        aliases = str(row.get("ALIASES", "")).strip()
        description = str(
            row.get("DESCRIPTION", "")
            or row.get("INFO", "")
            or row.get("COMMENT", "")
        ).strip()

        if not name:
            continue

        parts = [f"Продукт: {name}"]
        if aliases:
            parts.append(f"Аліаси: {aliases}")
        if description:
            parts.append(f"Опис: {description}")

        lines.append(" | ".join(parts))

    return "\n".join(lines)


# ================= CLEAN =================
def clean_and_structure(dialogue, replacements):
    if not dialogue:
        return None

    dictionary_text = "\n".join([f"{k} → {v}" for k, v in replacements.items()])

    prompt = f"""
Ти отримуєш транскрипт дзвінка після speech-to-text.

Словник замін є ОБОВʼЯЗКОВИМ.

Якщо слово або фраза присутні у словнику:
- використовуй ТІЛЬКИ варіант зі словника
- навіть якщо тобі здається, що інший варіант правильніший

ЗАБОРОНЕНО:
- змінювати слова, які є у словнику
- створювати власні варіанти, якщо є словник

Якщо є конфлікт між словником і твоєю логікою — пріоритет має словник.
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

def is_autoresponder(dialogue: str) -> bool:
    if not dialogue:
        return False

    text = dialogue.lower()

    triggers = [
        "залиште повідомлення",
        "після сигналу",
        "абонент недоступний",
        "не може відповісти",
        "voice mail",
        "voicemail",
        "please leave a message",
        "номер не обслуговується"
    ]

    return any(t in text for t in triggers)

# ================= GPT =================
def apply_defaults(features):
    defaults = {
        "manager_name_present": False,
        "manager_position_present": False,
        "company_present": False,
        "client_name_used": False,
        "purpose_present": False,
        "friendly_question": False,

        "bonus_offered": False,
        "bonus_has_type": False,
        "bonus_has_duration": False,

        "followup_type": "none",

        "objection_detected": False,
        "client_wants_to_end": False,
        "continuation_level": "none",
        "continuation_behavior": "neutral",

        "has_farewell": False,
        "is_limited_dialogue": False,

        "presentation_level": "none",
        "speech_quality": "bad",

        "assumption_made": False,

        "comment_match_level": "none",
        "comment_complete": False
    }

    for k, v in defaults.items():
        features.setdefault(k, v)

    return features


def extract_features_openai(dialogue, comment, kb_context=""):
    intro, middle, ending = extract_segments(dialogue)
    try:
        prompt = get_full_analysis_prompt_openai(intro, middle, ending, comment, kb_context)
    except TypeError:
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

        return apply_defaults(json.loads(match.group()))

    except Exception as e:
        st.error(f"GPT error: {e}")
        return {}


def extract_features_claude(dialogue, comment, kb_context=""):
    intro, middle, ending = extract_segments(dialogue)
    try:
        prompt = get_full_analysis_prompt_claude(intro, middle, ending, comment, kb_context)
    except TypeError:
        prompt = get_full_analysis_prompt(intro, middle, ending, comment)

    try:
        response = claude_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1000,
            messages=[
                {
                    "role": "user",
                    "content": f"Return ONLY valid JSON.\n{prompt}"
                }
            ]
        )

        text = response.content[0].text
        match = re.search(r"\{[\s\S]*\}", text)

        if not match:
            return {}

        return apply_defaults(json.loads(match.group()))

    except Exception as e:
        st.error(f"Claude error: {e}")
        return {}


# ================= SCORING =================
def score_call(f, meta, dialogue=None):
    s = {}

    # якщо автовідповідач → всі 0
    if dialogue and is_autoresponder(dialogue):
        return {
            "Встановлення контакту": 0,
            "Спроба презентації": 0,
            "Домовленість про наступний контакт": 0,
            "Пропозиція бонусу": 0,
            "Завершення розмови": 0,
            "Передзвон клієнту": 0,
            "Не додумувати": 0,
            "Якість мовлення": 0,
            "Професіоналізм": 0,
            "Оформлення картки": 0,
            "Утримання клієнта": 0,
            "Робота із запереченнями": 0
        }

    # ---------------- Контакт ----------------
    elements = sum([
    f["manager_name_present"],
    f["manager_position_present"],
    f["company_present"],
    f["client_name_used"],
    f["purpose_present"],
    f.get("friendly_question", False)
])

    s["Встановлення контакту"] = (
        7.5 if elements >= 4 else
        5 if elements == 3 else
        2.5 if elements == 2 else
        0
    )

    # ---------------- Спроба презентації ----------------
    level = f.get("presentation_level", "none")

    if level == "full":
        s["Спроба презентації"] = 5
    elif level == "partial":
        s["Спроба презентації"] = 2.5
    else:
        s["Спроба презентації"] = 0

    # ---------------- Домовленість ----------------
    fup = f.get("followup_type", "none")
    s["Домовленість про наступний контакт"] = (
        5 if fup == "exact_time"
        else 2.5 if fup == "offer"
        else 0
    )

    # ---------------- Бонус ----------------
    text = (dialogue or "").lower()
    
    bonus_keywords = [
        "бонус", "депозит", "вейджер", "відіграш",
        "48 год", "48год", "термін", "обмеження",
        "оберти", "фріспіни"
    ]
    
    found = [kw for kw in bonus_keywords if kw in text]
    
    cond = len(set(found))
    
    s["Пропозиція бонусу"] = (
        10 if cond >= 2 else
        5 if cond == 1 else
        0
    )

    # ---------------- Завершення ----------------
    s["Завершення розмови"] = 5 if f.get("has_farewell") else 0

    # ---------------- Передзвон ----------------
    repeat = meta["repeat_call"]
    
    if fup in ["none", "offer", "exact_time"]:
        s["Передзвон клієнту"] = 15
    else:
        s["Передзвон клієнту"] = (
            15 if repeat == "так, був протягом години"
            else 10 if repeat == "так, був протягом 2 годин"
            else 0
        )

    # ---------------- Не додумувати ----------------
    if f.get("assumption_made"):
        s["Не додумувати"] = 2.5
    else:
        s["Не додумувати"] = 5

    # ---------------- Якість мовлення ----------------
    quality = f.get("speech_quality", "bad")

    if quality == "good":
        s["Якість мовлення"] = 2.5
    else:
        s["Якість мовлення"] = 0

    # ---------------- Професіоналізм ----------------
    s["Професіоналізм"] = (
        5 if meta["bonus_check"] == "помилково нараховано" else 10
    )

    # ---------------- Картка ----------------
    match = f.get("comment_match_level", "none")
    complete = f.get("comment_complete", False)

    if match == "none":
        s["Оформлення картки"] = 0
    elif not complete:
        s["Оформлення картки"] = 2.5
    else:
        s["Оформлення картки"] = 5

    # ---------------- Утримання ----------------
    lvl = f.get("continuation_level", "none")

    if not f.get("client_wants_to_end"):
        behavior = f.get("continuation_behavior", "neutral")
        s["Утримання клієнта"] = (
            20 if behavior == "active"
            else 15 if behavior == "neutral"
            else 10 if behavior == "passive"
            else 0
        )
    else:
        s["Утримання клієнта"] = (
            20 if lvl == "strong"
            else 15 if lvl == "weak"
            else 10 if lvl == "formal"
            else 5 if lvl == "none"
            else 0
        )

    # ---------------- Заперечення ----------------
    if not f.get("objection_detected"):
        s["Робота із запереченнями"] = 10
    else:
        s["Робота із запереченнями"] = (
            10 if lvl == "strong"
            else 5 if lvl == "weak"
            else 0
        )

    return s


# ================= COMMENT =================
def generate_qa_comment(dialogue, scores):
    try:
        prompt = get_qa_comment_prompt(dialogue, scores)

        res = client.chat.completions.create(
            model="gpt-5.4",
            temperature=0.2,
            messages=[
                {"role": "system", "content": "Ти QA-аналітик дзвінків."},
                {"role": "user", "content": prompt}
            ]
        )

        return res.choices[0].message.content.strip()

    except Exception as e:
        st.error(f"Comment error: {e}")
        return "Помилка генерації коментаря"


def format_comment_for_sheet(comment):
    if not comment:
        return ""

    lines = [line.strip() for line in str(comment).splitlines() if line.strip()]
    return " | ".join(lines)

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
    kb_data = []
    kb_context = ""

    try:
        google_client = connect_google()
        dict_sheet = google_client.open_by_key(LOG_SHEET_ID).worksheet("DICT")
        replacements = load_replacements(dict_sheet)

        kb_sheet = google_client.open_by_key(KB_SHEET_ID).worksheet("INFO")
        kb_data = load_kb_data(kb_sheet)
        kb_context = build_kb_context(kb_data)
        
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

            # жорсткі заміни
            transcript = apply_replacements(transcript, replacements)
            clean_dialogue = clean_and_structure(transcript, replacements)
            clean_dialogue = apply_replacements(clean_dialogue, replacements)
            presentation_detected = detect_presentation(clean_dialogue, kb_data)

            if run_openai:
                features = extract_features_openai(
                    clean_dialogue,
                    call["manager_comment"],
                    kb_context
                )
            else:
                features = extract_features_claude(
                    clean_dialogue,
                    call["manager_comment"],
                    kb_context
                )
            
            # фільтр через базу знань
            if not presentation_detected:
                features["presentation_level"] = "none"

            if not features:
                st.warning("Помилка аналізу")
                continue

            scores = score_call(features, call, clean_dialogue)
            comment = generate_qa_comment(clean_dialogue, scores)
            comment_for_sheet = format_comment_for_sheet(comment)
            ai_label = "OpenAI" if run_openai else "Claude"

            st.session_state["results"].append({
                "scores": scores,
                "comment": comment
            })

            if google_client:
                try:
                    if not call["ret_sheet_id"]:
                        st.error("Не обрано проєкт або менеджера РЕТ")
                        continue

                    # 🟢 таблиця менеджера
                    sheet = google_client.open_by_key(call["ret_sheet_id"]).sheet1

                    # 🟢 формуємо оцінку одним рядком
                    total_score = sum(scores.values())

                    # 🟢 спочатку оцінки
                    res = write_to_google_sheet(sheet, call, scores) 
                    st.write("WRITE RESULT:", res)

                    # 🟢 запис у таблицю менеджера (твоя структура)
                    append_manager_log(
                        sheet,
                        call,
                        comment_for_sheet,
                        total_score,
                        ai_label
                    )

                    # 🟢 лог таблиця
                    log_sheet = google_client.open_by_key(LOG_SHEET_ID).worksheet("Лист 1")
                    append_qa_log(
                        log_sheet,
                        call,
                        transcript,
                        clean_dialogue,
                        comment,
                        total_score
                    )

                except Exception as e:
                    st.error(f"Google error: {e}")

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
        for line in res["comment"].split("\n"):     
            st.write(line)

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

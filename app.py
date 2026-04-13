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
    "Дар'я", "Надя", "Настя", "Владимира", "Діана", "Руслана", "Олексій", "Катерина"
]

FORBIDDEN_PROFESSIONALISM_PHRASES = [
    "Лотерея",
    "Акція",
    "Розіграш",
    "Реклама",
    "Подарунок",
    "Популяризація",
    "Лотерейний білет",
    "Даруємо",
    "Розігруємо",
    "Конкурс",
    "Кешбек",
    "Відшкодуємо",
    "Компенсація",
    "Повернення",
    "Фріспіни",
    "Безкоштовно",
    "Страхування",
    "страховка",
    "ставка без ризику",
    "фрібет",
    "Бездеп",
]

call_completion_statuses = [
    "⚪ (відсутній статус)",
    "🟢 (слухавку поклав клієнт)",
    "🟡 (технічні проблеми, зв'язок обірвався)",
    "🔴 (слухавку поклав менеджер)",
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
            repeat_col, completion_col = st.columns(2)
            with repeat_col:
                repeat_call = st.selectbox(
                    "Передзвон",
                    ["так, був протягом години", "так, був протягом 2 годин", "ні, не було"],
                    key=f"repeat_{idx}"
                )
            with completion_col:
                call_completion_status = st.selectbox(
                    "Завершення виклику",
                    call_completion_statuses,
                    key=f"call_completion_{idx}"
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
                "call_completion_status": call_completion_status,
                "manager_comment": manager_comment,
            })

# ================= TRANSCRIPTION =================
@st.cache_data(ttl=86400, show_spinner=False)
def transcribe_audio_cached(url):
    if not url:
        return {"ok": False, "error": "empty url", "transcript": None}

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
            return {"ok": False, "error": f"Deepgram error: {r.text}", "transcript": None}

        data = r.json()
        results = data.get("results", {})

        channels = results.get("channels", [])
        utterances = results.get("utterances", [])

        all_words = []

        if not channels and utterances:
            dialogue = []
            for u in utterances:
                speaker = f"ch_{u.get('speaker', 0)}"
                text = u.get("transcript", "")
                if text:
                    dialogue.append(f"{speaker}: {text}")
            return {"ok": True, "error": "", "transcript": "\n".join(dialogue)}

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
            return {"ok": False, "error": "Немає транскрипції", "transcript": None}

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

        return {"ok": True, "error": "", "transcript": "\n".join(dialogue)}

    except Exception as e:
        return {"ok": False, "error": f"Transcription exception: {str(e)}", "transcript": None}


def transcribe_audio(url):
    result = transcribe_audio_cached(url)
    if not result["ok"]:
        st.error(result["error"])
        return None
    return result["transcript"]


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
        "bonus_has_value": False,

        "followup_type": "none",

        "objection_detected": False,
        "client_wants_to_end": False,
        "continuation_level": "none",
        "continuation_behavior": "neutral",

        "has_farewell": False,
        "is_limited_dialogue": False,

        "presentation_level": "none",
        "speech_quality": "bad",
        "forbidden_words_used": False,
        "forbidden_words_detected": [],
        "conversation_logically_completed": False,
        "client_negative": False,
        "client_used_profanity": False,
        "manager_hung_up_before_client_finished": False,

        "assumption_made": False,

        "comment_match_level": "none",
        "comment_complete": False
    }

    for k, v in defaults.items():
        features.setdefault(k, v)

    return features


def build_dictionary_context(replacements):
    if not replacements:
        return "Словник замін не переданий."

    return "\n".join([f"{k} → {v}" for k, v in replacements.items()])


def normalize_forbidden_phrase(text):
    normalized = str(text or "").strip().lower()
    normalized = normalized.replace("’", "'").replace("`", "'").replace("ё", "е")
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized


def detect_forbidden_phrases_in_dialogue(dialogue):
    if not dialogue:
        return []

    manager_lines = []
    for line in str(dialogue).splitlines():
        stripped = line.strip()
        if stripped.startswith("Менеджер:") or stripped.startswith("ch_0:"):
            manager_lines.append(stripped.split(":", 1)[1].strip())

    manager_text = " ".join(manager_lines)
    if not manager_text:
        return []

    normalized_text = normalize_forbidden_phrase(manager_text)
    detected = []

    for phrase in FORBIDDEN_PROFESSIONALISM_PHRASES:
        normalized_phrase = normalize_forbidden_phrase(phrase)
        if not normalized_phrase:
            continue

        if " " in normalized_phrase:
            matched = normalized_phrase in normalized_text
        else:
            matched = re.search(rf"(?<!\w){re.escape(normalized_phrase)}(?!\w)", normalized_text) is not None

        if matched:
            detected.append(phrase)

    return detected


def validate_forbidden_words(features, dialogue):
    detected = detect_forbidden_phrases_in_dialogue(dialogue)
    features["forbidden_words_detected"] = detected
    features["forbidden_words_used"] = bool(detected)
    return features


def get_analysis_output_schema():
    return """
Поверни ONLY valid JSON такого формату:
{
  "cleaned_transcript": "очищений діалог",
  "qa_comment": "готовий QA-коментар по критеріях, кожен критерій з нового рядка",
  "features": {
    "manager_name_present": boolean,
    "manager_position_present": boolean,
    "company_present": boolean,
    "client_name_used": boolean,
    "purpose_present": boolean,
    "friendly_question": boolean,
    "presentation_level": "none" | "partial" | "full",
    "followup_type": "none" | "offer" | "exact_time",
    "bonus_offered": boolean,
    "bonus_has_type": boolean,
    "bonus_has_duration": boolean,
    "bonus_has_value": boolean,
    "has_farewell": boolean,
    "is_limited_dialogue": boolean,
    "objection_detected": boolean,
    "continuation_level": "none" | "formal" | "weak" | "strong" | "forced_end",
    "continuation_behavior": "active" | "neutral" | "passive" | "forced_end",
    "client_wants_to_end": boolean,
    "assumption_made": boolean,
    "comment_match_level": "none" | "partial" | "full",
    "comment_complete": boolean,
    "speech_quality": "bad" | "good",
    "forbidden_words_used": boolean,
    "forbidden_words_detected": ["рядок 1", "рядок 2"],
    "conversation_logically_completed": boolean,
    "client_negative": boolean,
    "client_used_profanity": boolean,
    "manager_hung_up_before_client_finished": boolean
  }
}
"""


def build_combined_analysis_prompt(prompt_body, raw_dialogue, replacements):
    dictionary_context = build_dictionary_context(replacements)
    forbidden_words_list = ", ".join(FORBIDDEN_PROFESSIONALISM_PHRASES)
    return f"""
{prompt_body}

---------------------
РЎР›РћР’РќРРљ Р—РђРњР†Рќ
---------------------

РЎР»РѕРІРЅРёРє Р·Р°РјС–РЅ С” РћР‘РћР’'РЇР—РљРћР’РРњ.
Якщо слово або фраза є у словнику, використовуй тільки варіант зі словника.
Не вигадуй власних варіантів, якщо слово є у словнику.

{dictionary_context}

---------------------
РћР§РРЎРўРљРђ РўР РђРќРЎРљР РРџРўРЈ
---------------------

Спочатку очисти транскрипт:
- виправ помилки розпізнавання
- застосуй словник замін
- не змінюй сенс
- не скорочуй текст
- заміни ch_0 на "Менеджер", ch_1 на "Клієнт"

Після цього:
- проаналізуй вже очищений транскрипт
- сформуй готовий qa_comment у тому ж запиті
- qa_comment має бути українською, по одному критерію на рядок
- для критерію "Професіоналізм" перевір лише репліки менеджера
- якщо менеджер вжив хоча б одне заборонене слово або фразу, поверни "forbidden_words_used": true
- у "forbidden_words_detected" поверни точні слова або фрази, які вжив менеджер
- якщо "forbidden_words_used": true, критерій "Професіоналізм" має оцінюватися в 0 балів
- якщо "forbidden_words_used": true, у qa_comment ОБОВ'ЯЗКОВО додай окремий рядок про критерій "Професіоналізм" і вкажи конкретне заборонене слово або фразу
- додатково визнач:
  "conversation_logically_completed" = true, якщо розмова по суті завершена
  "client_negative" = true, якщо клієнт проявляє негатив
  "client_used_profanity" = true, якщо клієнт використовує нецензурну лексику
  "manager_hung_up_before_client_finished" = true, якщо менеджер не дослухав клієнта і сам завершив незавершену розмову

Заборонені слова і фрази для критерію "Професіоналізм":
{forbidden_words_list}

{get_analysis_output_schema()}

РЎРР РР™ РўР РђРќРЎРљР РРџРў:
{raw_dialogue}
"""


def parse_analysis_response(text):
    match = re.search(r"\{[\s\S]*\}", text or "")
    if not match:
        return None

    payload = json.loads(match.group())
    features = apply_defaults(payload.get("features", {}))

    return {
        "cleaned_transcript": (payload.get("cleaned_transcript") or "").strip(),
        "qa_comment": (payload.get("qa_comment") or "").strip(),
        "features": features,
    }


def extract_features_openai(dialogue, comment, kb_context="", replacements=None):
    intro, middle, ending = extract_segments(dialogue)
    try:
        base_prompt = get_full_analysis_prompt_openai(intro, middle, ending, comment, kb_context)
    except TypeError:
        base_prompt = get_full_analysis_prompt(intro, middle, ending, comment)

    prompt = build_combined_analysis_prompt(base_prompt, dialogue, replacements or {})

    try:
        res = client.chat.completions.create(
            model="gpt-5.4",
            temperature=0,
            messages=[
                {"role": "system", "content": "Return only valid JSON."},
                {"role": "user", "content": prompt}
            ]
        )
        parsed = parse_analysis_response(res.choices[0].message.content)
        return parsed or {}

    except Exception as e:
        st.error(f"GPT error: {e}")
        return {}


def extract_features_claude(dialogue, comment, kb_context="", replacements=None):
    intro, middle, ending = extract_segments(dialogue)
    try:
        base_prompt = get_full_analysis_prompt_claude(intro, middle, ending, comment, kb_context)
    except TypeError:
        base_prompt = get_full_analysis_prompt(intro, middle, ending, comment)

    prompt = build_combined_analysis_prompt(base_prompt, dialogue, replacements or {})

    try:
        response = claude_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            messages=[
                {
                    "role": "user",
                    "content": f"Return ONLY valid JSON.\n{prompt}"
                }
            ]
        )

        parsed = parse_analysis_response(response.content[0].text)
        return parsed or {}

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
    if not f.get("bonus_offered"):
        s["Пропозиція бонусу"] = 0
    else:
        bonus_conditions = sum([
            bool(f.get("bonus_has_type")),
            bool(f.get("bonus_has_duration")),
            bool(f.get("bonus_has_value"))
        ])
        s["Пропозиція бонусу"] = 10 if bonus_conditions >= 2 else 5

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
    if f.get("forbidden_words_used"):
        s["Професіоналізм"] = 0
    else:
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

    return apply_call_completion_rules(s, f, meta)


def format_comment_for_sheet(comment):
    if not comment:
        return ""

    lines = [line.strip() for line in str(comment).splitlines() if line.strip()]
    return " | ".join(lines)


def build_readable_qa_comment(features, scores, call):
    lines = []

    contact_elements = sum([
        bool(features.get("manager_name_present")),
        bool(features.get("manager_position_present")),
        bool(features.get("company_present")),
        bool(features.get("client_name_used")),
        bool(features.get("purpose_present")),
    ])
    if scores.get("Встановлення контакту", 0) >= 5:
        lines.append("Встановлення контакту: менеджер коректно представився, звернувся до клієнта та озвучив мету дзвінка.")
    elif contact_elements >= 2:
        lines.append("Встановлення контакту: контакт встановлено частково, але не всі обов'язкові елементи були озвучені.")
    else:
        lines.append("Встановлення контакту: менеджер не представився повноцінно і не окреслив мету дзвінка.")

    presentation_level = features.get("presentation_level", "none")
    if presentation_level == "full":
        lines.append("Спроба презентації: менеджер презентував продукт, пояснив суть пропозиції та що клієнту потрібно зробити.")
    elif presentation_level == "partial":
        lines.append("Спроба презентації: менеджер лише коротко згадав продукт або активність без повного пояснення суті та дії для клієнта.")
    else:
        lines.append("Спроба презентації: презентації продукту не було; інформація лише про бонус не рахується як презентація.")

    followup_type = features.get("followup_type", "none")
    if followup_type == "exact_time":
        lines.append("Домовленість про наступний контакт: узгоджено конкретний час наступного дзвінка.")
    elif followup_type == "offer":
        lines.append("Домовленість про наступний контакт: передзвон запропоновано, але без узгодженого точного часу.")
    else:
        lines.append("Домовленість про наступний контакт: домовленості про наступний дзвінок не було.")

    if not features.get("bonus_offered"):
        lines.append("Пропозиція бонусу: бонус клієнту не озвучено.")
    else:
        bonus_details = []
        if features.get("bonus_has_type"):
            bonus_details.append("тип бонусу")
        if features.get("bonus_has_duration"):
            bonus_details.append("термін дії")
        if features.get("bonus_has_value"):
            bonus_details.append("розмір бонусу")
        if scores.get("Пропозиція бонусу", 0) >= 10:
            lines.append("Пропозиція бонусу: бонус озвучено як вигоду, названо щонайменше дві його умови.")
        else:
            detail_text = ", ".join(bonus_details) if bonus_details else "лише частину умов"
            lines.append(f"Пропозиція бонусу: бонус згадано формально, озвучено {detail_text}.")

    if features.get("has_farewell"):
        lines.append("Завершення розмови: розмову завершено з прощанням.")
    else:
        lines.append("Завершення розмови: прощання наприкінці розмови відсутнє.")

    repeat_call = call.get("repeat_call", "")
    if scores.get("Передзвон клієнту", 0) == 15:
        if repeat_call == "так, був протягом години":
            lines.append("Передзвон клієнту: передзвон виконано протягом години.")
        else:
            lines.append("Передзвон клієнту: штрафу немає, додатковий передзвон у цьому сценарії не був потрібний.")
    elif scores.get("Передзвон клієнту", 0) == 10:
        lines.append("Передзвон клієнту: передзвон був, але не одразу, а протягом двох годин.")
    else:
        lines.append("Передзвон клієнту: потрібного передзвону не було, тому критерій не виконано.")

    if features.get("assumption_made"):
        lines.append("Не додумувати: менеджер припускав або додумував інформацію замість опори на факти з діалогу.")
    else:
        lines.append("Не додумувати: менеджер не додумував зайвого і тримався фактів розмови.")

    if features.get("speech_quality") == "good":
        lines.append("Якість мовлення: мовлення достатньо чисте та зрозуміле для аналізу.")
    else:
        lines.append("Якість мовлення: у мовленні є проблеми, які заважають сприйняттю або точному аналізу.")

    detected = [
        str(item).strip()
        for item in features.get("forbidden_words_detected", [])
        if str(item).strip()
    ]
    if detected:
        lines.append(
            "Професіоналізм: 0 балів, менеджер використав заборонені слова/фрази: "
            f"{', '.join(detected)}."
        )
    elif call.get("bonus_check") == "помилково нараховано":
        lines.append("Професіоналізм: критерій знижено через помилково нарахований бонус.")
    else:
        lines.append("Професіоналізм: заборонених слів зі списку не виявлено.")

    comment_match_level = features.get("comment_match_level", "none")
    if comment_match_level == "full" and features.get("comment_complete"):
        lines.append("Оформлення картки: коментар у картці відповідає змісту дзвінка і містить ключову інформацію.")
    elif comment_match_level == "partial":
        lines.append("Оформлення картки: коментар у картці неповний або не повністю відповідає змісту розмови.")
    else:
        lines.append("Оформлення картки: коментар у картці відсутній або не відображає результат дзвінка.")

    if not features.get("objection_detected"):
        lines.append("Робота із запереченнями: заперечень від клієнта не було.")
    elif scores.get("Робота із запереченнями", 0) >= 10:
        lines.append("Робота із запереченнями: менеджер відпрацював заперечення аргументовано.")
    elif scores.get("Робота із запереченнями", 0) >= 5:
        lines.append("Робота із запереченнями: була спроба відпрацювати заперечення, але недостатньо глибока.")
    else:
        lines.append("Робота із запереченнями: заперечення не були відпрацьовані.")

    if features.get("client_wants_to_end"):
        continuation_level = features.get("continuation_level", "none")
        if scores.get("Утримання клієнта", 0) >= 20 or continuation_level == "strong":
            lines.append("Утримання клієнта: менеджер зробив кілька змістовних спроб втримати клієнта в розмові.")
        elif scores.get("Утримання клієнта", 0) >= 15 or continuation_level == "weak":
            lines.append("Утримання клієнта: була одна реальна спроба втримати клієнта в розмові.")
        elif scores.get("Утримання клієнта", 0) >= 10 or continuation_level == "formal":
            lines.append("Утримання клієнта: реакція менеджера була формальною, без повноцінного утримання.")
        else:
            lines.append("Утримання клієнта: менеджер не втримував клієнта в розмові, коли це було потрібно.")
    else:
        continuation_behavior = features.get("continuation_behavior", "neutral")
        if scores.get("Утримання клієнта", 0) >= 20 or continuation_behavior == "active":
            lines.append("Утримання клієнта: менеджер активно вів діалог і не давав розмові згаснути.")
        elif scores.get("Утримання клієнта", 0) >= 15 or continuation_behavior == "neutral":
            lines.append("Утримання клієнта: менеджер підтримував розмову на нормальному рівні без провалів.")
        elif scores.get("Утримання клієнта", 0) >= 10 or continuation_behavior == "passive":
            lines.append("Утримання клієнта: розмову вели пасивно, без достатньої ініціативи з боку менеджера.")
        else:
            lines.append("Утримання клієнта: менеджер допустив втрату розмови або сам спровокував її завершення.")

    return "\n".join(lines)


def use_test_project_scores_sheet(call):
    return (
        call.get("project") == "ТЕСТ"
        and call.get("ret_manager") in {"Жарікова Анастасія", "Бурий Андрій"}
    )


def get_manager_sheet_settings(call):
    if use_test_project_scores_sheet(call):
        return {
            "worksheet_name": "Оцінки",
            "start_column": 4,
            "scores_start_row": 1,
            "criteria_start_row": 5,
            "log_start_row": 20,
        }

    return {
        "worksheet_name": "Оцінки",
        "start_column": 4,
        "scores_start_row": 88,
        "criteria_start_row": 93,
        "log_start_row": 110,
    }


def apply_call_completion_rules(scores, features, meta):
    status = meta.get("call_completion_status", "")
    immediate_repeat = meta.get("repeat_call") == "так, був протягом години"
    has_any_repeat = meta.get("repeat_call") in {
        "так, був протягом години",
        "так, був протягом 2 годин",
    }
    logical_completion = bool(features.get("conversation_logically_completed"))
    has_farewell = bool(features.get("has_farewell"))
    bonus_offered = bool(features.get("bonus_offered"))
    has_followup = features.get("followup_type", "none") != "none"
    client_negative = bool(features.get("client_negative"))
    client_used_profanity = bool(features.get("client_used_profanity"))
    manager_hung_up_early = bool(features.get("manager_hung_up_before_client_finished"))

    if logical_completion and has_farewell:
        return scores

    if status == "🟢 (слухавку поклав клієнт)":
        if (
            not logical_completion
            and not has_farewell
            and bonus_offered
            and has_followup
            and immediate_repeat
        ):
            return scores

        if client_negative and not client_used_profanity:
            if not immediate_repeat:
                scores["Передзвон клієнту"] = 0
            return scores

        if client_negative and client_used_profanity and not immediate_repeat:
            return scores

        if (
            not logical_completion
            and not has_farewell
            and not bonus_offered
            and not has_followup
            and not immediate_repeat
        ):
            scores["Утримання клієнта"] = 0
            return scores

    if status == "🔴 (слухавку поклав менеджер)":
        if manager_hung_up_early or client_negative:
            scores["Утримання клієнта"] = 0
            return scores

    if status == "🟡 (технічні проблеми, зв'язок обірвався)":
        if not logical_completion and not has_any_repeat:
            scores["Передзвон клієнту"] = 0
            return scores

    return scores

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

            transcript = apply_replacements(transcript, replacements)

            if run_openai:
                analysis_result = extract_features_openai(
                    transcript,
                    call["manager_comment"],
                    kb_context,
                    replacements
                )
            else:
                analysis_result = extract_features_claude(
                    transcript,
                    call["manager_comment"],
                    kb_context,
                    replacements
                )

            if not analysis_result:
                st.warning("Помилка аналізу")
                continue

            clean_dialogue = analysis_result.get("cleaned_transcript") or transcript
            clean_dialogue = apply_replacements(clean_dialogue, replacements)
            features = analysis_result.get("features", {})
            features = validate_forbidden_words(features, clean_dialogue)
            comment = analysis_result.get("qa_comment", "").strip()
            presentation_detected = detect_presentation(clean_dialogue, kb_data)

            # фільтр через базу знань
            if not presentation_detected:
                features["presentation_level"] = "none"

            if not features:
                st.warning("Помилка аналізу")
                continue

            scores = score_call(features, call, clean_dialogue)
            comment = build_readable_qa_comment(features, scores, call)
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
                    workbook = google_client.open_by_key(call["ret_sheet_id"])
                    sheet_settings = get_manager_sheet_settings(call)
                    scores_sheet = (
                        workbook.worksheet(sheet_settings["worksheet_name"])
                        if sheet_settings["worksheet_name"]
                        else workbook.sheet1
                    )
                    scores_start_column = sheet_settings["start_column"]

                    # 🟢 формуємо оцінку одним рядком
                    total_score = sum(scores.values())

                    # 🟢 спочатку оцінки
                    res = write_to_google_sheet(
                        scores_sheet,
                        call,
                        scores,
                        start_column=scores_start_column,
                        start_row=sheet_settings["scores_start_row"],
                        criteria_start_row=sheet_settings["criteria_start_row"],
                    )
                    st.write("WRITE RESULT:", res)

                    # 🟢 запис у таблицю менеджера (твоя структура)
                    append_manager_log(
                        scores_sheet,
                        call,
                        comment_for_sheet,
                        total_score,
                        ai_label,
                        start_row=sheet_settings["log_start_row"],
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

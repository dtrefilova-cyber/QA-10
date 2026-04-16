import streamlit as st
import pandas as pd
import requests
import json
import re
from google_sheets import (
    append_log_info,
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
ANALYSIS_CACHE_VERSION = "2026-04-16-1"

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

@st.cache_data(ttl=300, show_spinner=False)
def get_managers_config():
    google_client = connect_google()
    return load_managers_config(google_client, LOG_SHEET_ID)


@st.cache_data(ttl=300, show_spinner=False)
def get_reference_data():
    google_client = connect_google()
    dict_sheet = google_client.open_by_key(LOG_SHEET_ID).worksheet("DICT")
    replacements = load_replacements(dict_sheet)

    kb_sheet = google_client.open_by_key(KB_SHEET_ID).worksheet("INFO")
    kb_data = load_kb_data(kb_sheet)
    kb_context = build_kb_context(kb_data)
    return replacements, kb_data, kb_context


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
        pattern = re.compile(rf"(?<!\w){re.escape(k)}(?!\w)", re.IGNORECASE)
        if pattern.search(text):
            text = pattern.sub(v, text)

    return text

def detect_presentation(dialogue, kb_data):
    if not dialogue:
        return False

    manager_lines = []
    for line in str(dialogue).splitlines():
        stripped = line.strip()
        if stripped.startswith("Менеджер:") or stripped.startswith("ch_0:"):
            manager_lines.append(stripped.split(":", 1)[1].strip())

    text = " ".join(manager_lines).lower()
    if not text:
        return False

    for row in kb_data:
        name = (row.get("NAME") or "").lower()
        aliases = (row.get("ALIASES") or "").lower().split(";")

        variants = [name] + aliases

        for v in variants:
            v = v.strip()
            if v and v in text:
                return True

    return False


def extract_role_lines(dialogue):
    manager_lines = []
    client_lines = []

    for raw_line in str(dialogue or "").splitlines():
        stripped = raw_line.strip()
        if ":" not in stripped:
            continue

        speaker, text = stripped.split(":", 1)
        speaker = speaker.strip().lower()
        text = text.strip()
        if not text:
            continue

        if speaker in {"менеджер", "ch_0"}:
            manager_lines.append(text)
        elif speaker in {"клієнт", "клиент", "ch_1"}:
            client_lines.append(text)

    return manager_lines, client_lines


def has_any_marker(text, markers):
    normalized = f" {str(text or '').lower()} "
    return any(marker in normalized for marker in markers)


def normalize_presentation_level(features, dialogue, kb_data):
    manager_lines, _ = extract_role_lines(dialogue)
    manager_text = " ".join(manager_lines).lower()

    if not manager_text:
        return features

    level = features.get("presentation_level", "none")
    has_bonus_word = "бонус" in manager_text
    has_product_mention = detect_presentation(dialogue, kb_data)

    loyalty_markers = [
        "програм",
        "лояльн",
        "монет",
        "медал",
        "спін",
        "спини",
        "підбірк",
        "добірк",
        "активн",
        "новинк",
        "продукт",
        "слот",
    ]
    location_markers = [
        "на сайті",
        "в додатку",
        "у додатку",
        "в особистому кабінеті",
        "в особистому",
        "у розділі",
        "в розділі",
        "знайдете",
        "можна знайти",
        "де знайти",
        "на головній",
    ]
    sent_markers = [
        "надішлю",
        "відправлю",
        "скину",
        "на пошту",
        "у смс",
        "в смс",
        "в вайбер",
        "у вайбер",
        "в телеграм",
        "у телеграм",
    ]
    explanation_markers = [
        "це ",
        "там ",
        "є ",
        "зможете",
        "потрібно",
        "треба",
    ]

    has_loyalty_mention = has_any_marker(manager_text, loyalty_markers)
    has_location = has_any_marker(manager_text, location_markers)
    has_sent_info = has_any_marker(manager_text, sent_markers)
    has_explanation = has_any_marker(manager_text, explanation_markers)

    bonus_only = has_bonus_word and not (has_product_mention or has_loyalty_mention)
    if bonus_only:
        features["presentation_level"] = "none"
        return features

    if has_product_mention or has_loyalty_mention or has_location or has_sent_info:
        if has_location and (has_product_mention or has_loyalty_mention):
            features["presentation_level"] = "full"
        elif level == "none":
            features["presentation_level"] = "partial"
        elif level not in {"partial", "full"}:
            features["presentation_level"] = "partial"

        if (
            features["presentation_level"] == "partial"
            and (has_sent_info and (has_product_mention or has_loyalty_mention or has_explanation))
        ):
            features["presentation_level"] = "full"

    return features


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
        "noise_reaction": "none",

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
        "followup_attempts_count": 0,
        "client_hung_up_interrupted": False,
        "client_sick": False,
        "manager_wished_recovery": False,
        "client_military": False,
        "manager_thanked_for_service": False,
        "client_driving_or_no_phone": False,
        "client_not_actual_client": False,
        "manager_shared_bonus_with_third_party": False,
        "client_unethical_behavior": False,
        "manager_unethical_response": False,

        "comment_match_level": "none",
        "comment_complete": False,
        "card_has_reason": False,
        "card_has_followup_time": False
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


def validate_assumption_made(features, dialogue):
    manager_lines, client_lines = extract_role_lines(dialogue)
    manager_text = " ".join(manager_lines).lower()
    client_text = " ".join(client_lines).lower()
    if not manager_text:
        features["assumption_made"] = False
        return features

    assumption_markers = [
        "вам зараз незручно",
        "чи зручно говорити",
        "я вам не заважаю",
        "давайте іншим разом",
        "ви, мабуть, зайняті",
        "ви мабуть зайняті",
        "немає часу так спілкуватися",
        "вам, мабуть, нецікаво",
        "вам, мабуть, незручно",
        "вам, мабуть, не до розмови",
        "вам мабуть нецікаво",
        "вам мабуть незручно",
        "вам мабуть не до розмови",
        "вам незручно",
        "не дуже вчасно набрав",
        "ви зайняті",
    ]

    client_state_markers = [
        "я зайнятий",
        "я занята",
        "мені незручно",
        "не можу говорити",
        "я за кермом",
        "передзвоніть",
    ]

    if any(marker in manager_text for marker in assumption_markers):
        if any(marker in client_text for marker in client_state_markers):
            features["assumption_made"] = False
        else:
            features["assumption_made"] = True
    else:
        features["assumption_made"] = False

    return features


def validate_bonus_features(features, dialogue):
    manager_lines, _ = extract_role_lines(dialogue)
    manager_text = " ".join(manager_lines).lower()

    if not manager_text:
        return features

    offer_markers = [
        "нарах",
        "дам бонус",
        "буде бонус",
        "будуть бонус",
        "залиш",
        "доступн",
        "від менеджера",
        "подарую",
        "отримаєте",
        "отримаєш",
        "можу дати",
    ]
    type_markers = [
        "фс",
        "fs",
        "фріспін",
        "фриспін",
        "спін",
        "спини",
        "кешбек",
        "кешбеку",
        "бонус на депозит",
        "бездеп",
        "фрібет",
        "від менеджера",
    ]
    duration_markers = [
        "годин",
        "днів",
        "день",
        "тиж",
        "до кінця",
        "сьогодні",
        "завтра",
        "48",
        "24",
        "термін дії",
        "діє",
    ]
    value_markers = [
        "%",
        "відсот",
        "грн",
        "грив",
        "сума",
        "депозит",
        "поповнен",
        "вейдж",
        "відіграш",
        "x",
        "х",
        "на 700",
        "на 1000",
        "50",
        "70",
        "100",
        "200",
    ]

    has_offer = has_any_marker(manager_text, offer_markers)
    if not has_offer:
        features["bonus_offered"] = False
        features["bonus_has_type"] = False
        features["bonus_has_duration"] = False
        features["bonus_has_value"] = False
        return features

    features["bonus_offered"] = True
    features["bonus_has_type"] = has_any_marker(manager_text, type_markers)
    features["bonus_has_duration"] = has_any_marker(manager_text, duration_markers)
    features["bonus_has_value"] = has_any_marker(manager_text, value_markers)
    return features


def validate_professionalism_features(features, dialogue):
    manager_lines, client_lines = extract_role_lines(dialogue)
    manager_text = " ".join(manager_lines).lower()
    client_text = " ".join(client_lines).lower()

    direct_client_markers = [
        "ви",
        "вам",
        "вас",
        "з вами",
    ]
    third_party_markers = [
        "це не",
        "його немає",
        "її немає",
        "мама",
        "тато",
        "дружина",
        "чоловік",
        "син",
        "донька",
        "дочка",
        "брат",
        "сестра",
        "подруга",
    ]

    has_direct_client_communication = has_any_marker(manager_text, direct_client_markers)
    has_clear_third_party_context = has_any_marker(client_text, third_party_markers) or has_any_marker(manager_text, third_party_markers)

    if has_direct_client_communication or not has_clear_third_party_context:
        features["client_not_actual_client"] = False

    return features


def validate_objection_and_retention(features, dialogue):
    manager_lines, client_lines = extract_role_lines(dialogue)
    manager_text = " ".join(manager_lines).lower()
    client_text = " ".join(client_lines).lower()

    end_call_markers = [
        "не можу говорити",
        "не можу зараз",
        "немає часу говорити",
        "я зайнятий",
        "я занята",
        "передзвоніть",
        "за кермом",
        "незручно говорити",
    ]
    product_objection_markers = [
        "не хочу грати",
        "не цікаво",
        "нецікаво",
        "не хочу бонус",
        "не хочу продукт",
        "не потрібно",
        "не треба",
        "не буду грати",
        "не хочу",
    ]
    real_retention_markers = [
        "буквально хвилин",
        "буквально секунд",
        "1 хвилин",
        "одну хвилин",
        "дуже коротко",
        "скажу головне",
        "лише головне",
        "одразу головне",
        "коротко поясню",
        "коротко розкажу",
    ]
    callback_only_markers = [
        "коли вам передзвонити",
        "на який час",
        "о котрій",
        "коли буде зручно",
    ]

    client_wants_to_end = has_any_marker(client_text, end_call_markers)
    product_objection = has_any_marker(client_text, product_objection_markers)
    real_retention = has_any_marker(manager_text, real_retention_markers)
    callback_only = has_any_marker(manager_text, callback_only_markers)
    bonus_only = "бонус" in manager_text and not real_retention

    if client_wants_to_end:
        features["client_wants_to_end"] = True

    if client_wants_to_end and not product_objection:
        features["objection_detected"] = False

    if product_objection:
        features["objection_detected"] = True
        if features.get("continuation_level") == "none" and (real_retention or len(manager_lines) > 0):
            features["continuation_level"] = "weak"

    if client_wants_to_end:
        if real_retention:
            if features.get("continuation_level") not in {"strong", "weak"}:
                features["continuation_level"] = "weak"
        elif callback_only or bonus_only:
            if features.get("continuation_level") == "strong":
                features["continuation_level"] = "formal"
            elif features.get("continuation_level") == "none":
                features["continuation_level"] = "formal"

    if features.get("continuation_level") == "strong":
        strong_count = 0
        strong_count += int(real_retention)
        strong_count += int("інший раз" in manager_text or "через годину" in manager_text or "ближче до вечора" in manager_text)
        if strong_count < 2:
            features["continuation_level"] = "weak"

    if features.get("objection_detected") and features.get("continuation_level") == "none":
        if features.get("client_hung_up_interrupted") or len(manager_lines) > 0:
            features["continuation_level"] = "weak"

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
      "noise_reaction": "none" | "correct" | "incorrect",
      "presentation_level": "none" | "partial" | "full",
      "followup_type": "none" | "offer" | "exact_time",
      "followup_attempts_count": integer,
      "client_hung_up_interrupted": boolean,
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
      "client_sick": boolean,
      "manager_wished_recovery": boolean,
      "client_military": boolean,
      "manager_thanked_for_service": boolean,
      "client_driving_or_no_phone": boolean,
      "client_not_actual_client": boolean,
      "manager_shared_bonus_with_third_party": boolean,
      "client_unethical_behavior": boolean,
      "manager_unethical_response": boolean,
      "comment_match_level": "none" | "partial" | "full",
      "comment_complete": boolean,
      "card_has_reason": boolean,
      "card_has_followup_time": boolean,
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
    return f"""
{prompt_body}

---------------------
CLEAN
---------------------

- очисти транскрипт без зміни сенсу
- застосуй словник замін
- не скорочуй текст
- заміни ch_0 на "Менеджер", ch_1 на "Клієнт"
- словник використовуй тільки для очистки, не для оцінювання

{dictionary_context}

---------------------
ANALYSIS
---------------------

- аналізуй тільки очищений транскрипт
- поверни тільки `features`, `cleaned_transcript` і `qa_comment`
- `qa_comment` формуй лише з фактів і значень `features`, без загального враження
- якщо `presentation_level` не `none`, коментар про презентацію має це відображати
- бонус сам по собі не є презентацією
- слово "бонус" без реальних умов не робить `bonus_has_type`, `bonus_has_duration`, `bonus_has_value` істинними
- домовленість про передзвін, питання про час і проста згадка бонусу не є утриманням
- якщо є лише одна спроба утримання, не пиши про кілька спроб
- "не хочу говорити" / "я зайнятий" / "передзвоніть" = утримання, не заперечення
- "не хочу продукт / гру / бонус" = заперечення
- `client_not_actual_client` став тільки якщо прямої комунікації з клієнтом немає
- додатково визнач:
  "conversation_logically_completed" = true, якщо розмова по суті завершена
  "client_negative" = true, якщо клієнт проявляє негатив
  "client_used_profanity" = true, якщо клієнт використовує нецензурну лексику
  "manager_hung_up_before_client_finished" = true, якщо менеджер не дослухав клієнта і сам завершив незавершену розмову

{get_analysis_output_schema()}

СИРИЙ ТРАНСКРИПТ:
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


@st.cache_data(show_spinner=False)
def analyze_call_cached(ai_provider, url, call_date, dialogue, manager_comment, kb_context, replacements, cache_version):
    if ai_provider == "openai":
        return extract_features_openai(
            dialogue,
            manager_comment,
            kb_context,
            replacements,
        )

    return extract_features_claude(
        dialogue,
        manager_comment,
        kb_context,
        replacements,
    )


# ================= SCORING =================
def score_call(f, meta, dialogue=None):
    s = {}
    noise_reaction = f.get("noise_reaction", "none")
    followup_type = f.get("followup_type", "none")
    followup_attempts_count = int(f.get("followup_attempts_count") or 0)
    is_military_client = bool(f.get("client_military"))
    is_driving_or_no_phone = bool(f.get("client_driving_or_no_phone"))
    unethical_client_behavior = bool(f.get("client_unethical_behavior"))
    manager_unethical_response = bool(f.get("manager_unethical_response"))

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

    # Обмежений діалог (клієнт зайнятий/за кермом/просить передзвонити тощо):
    # за правилами промпта не занижуємо за відсутність презентації/аргументації
    # та не штрафуємо за відсутність утримання.
    limited_dialogue = bool(f.get("is_limited_dialogue"))

    # ---------------- Контакт ----------------
    elements = sum([
        f["manager_name_present"],
        f["manager_position_present"],
        f["company_present"],
        f["client_name_used"],
        f["purpose_present"],
        f.get("friendly_question", False) or noise_reaction == "correct"
    ])

    contact_score = (
        7.5 if elements >= 4 else
        5 if elements == 3 else
        2.5 if elements == 2 else
        0
    )

    if not f.get("client_name_used"):
        contact_score -= 2.5

    if (
        (f.get("client_sick") and not f.get("manager_wished_recovery"))
        or (is_military_client and not f.get("manager_thanked_for_service"))
    ):
        contact_score -= 2.5

    s["Встановлення контакту"] = max(0, contact_score)

    # ---------------- Спроба презентації ----------------
    level = f.get("presentation_level", "none")

    if is_driving_or_no_phone:
        s["Спроба презентації"] = 5
    elif level == "full":
        s["Спроба презентації"] = 5
    elif level == "partial":
        s["Спроба презентації"] = 2.5
    else:
        s["Спроба презентації"] = 0

    if limited_dialogue:
        s["Спроба презентації"] = 5

    # ---------------- Домовленість ----------------
    s["Домовленість про наступний контакт"] = (
        5 if (
            followup_type == "exact_time"
            or followup_attempts_count >= 2
            or (
                meta.get("call_completion_status") == "🟢 (слухавку поклав клієнт)"
                and f.get("client_hung_up_interrupted")
            )
        )
        else 2.5 if followup_type == "offer"
        else 0
    )

    # ---------------- Бонус ----------------
    if is_driving_or_no_phone:
        s["Пропозиція бонусу"] = 10
    elif not f.get("bonus_offered"):
        s["Пропозиція бонусу"] = 0
    else:
        bonus_conditions = sum([
            bool(f.get("bonus_has_type")),
            bool(f.get("bonus_has_duration")),
            bool(f.get("bonus_has_value"))
        ])
        if bonus_conditions <= 0:
            s["Пропозиція бонусу"] = 0
        else:
            s["Пропозиція бонусу"] = 10 if bonus_conditions >= 2 else 5

    # ---------------- Завершення ----------------
    s["Завершення розмови"] = 5 if f.get("has_farewell") else 0

    # ---------------- Передзвон ----------------
    repeat = meta["repeat_call"]

    if followup_type == "none":
        s["Передзвон клієнту"] = 15
    else:
        s["Передзвон клієнту"] = (
            15 if repeat == "так, був протягом години"
            else 10 if repeat == "так, був протягом 2 годин"
            else 0
        )

    # ---------------- Не додумувати ----------------
    if f.get("assumption_made"):
        s["Не додумувати"] = 0
    else:
        s["Не додумувати"] = 5

    # ---------------- Якість мовлення ----------------
    quality = f.get("speech_quality", "bad")

    if quality == "good":
        s["Якість мовлення"] = 2.5
    else:
        s["Якість мовлення"] = 0

    # ---------------- Професіоналізм ----------------
    if f.get("forbidden_words_used") or (
        f.get("client_not_actual_client") and f.get("manager_shared_bonus_with_third_party")
    ):
        s["Професіоналізм"] = 0
    else:
        s["Професіоналізм"] = (
            5 if meta["bonus_check"] == "помилково нараховано" else 10
        )

    # ---------------- Картка ----------------
    card_elements = sum([
        bool(f.get("card_has_reason")),
        bool(f.get("card_has_followup_time")),
    ])
    s["Оформлення картки"] = 5 if card_elements == 2 else 2.5 if card_elements == 1 else 0

    # ---------------- Утримання ----------------
    lvl = f.get("continuation_level", "none")

    if is_military_client:
        s["Утримання клієнта"] = 20
    elif limited_dialogue:
        s["Утримання клієнта"] = 20
    elif not f.get("client_wants_to_end"):
        behavior = f.get("continuation_behavior", "neutral")
        s["Утримання клієнта"] = (
            20 if behavior == "active"
            else 15 if behavior == "neutral"
            else 0 if behavior == "passive"
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
    if is_military_client:
        s["Робота із запереченнями"] = 10
    elif limited_dialogue:
        s["Робота із запереченнями"] = 10
    elif not f.get("objection_detected"):
        s["Робота із запереченнями"] = 10
    else:
        s["Робота із запереченнями"] = (
            10 if lvl == "strong"
            else 5 if lvl in {"weak", "formal"}
            else 0
        )

    if unethical_client_behavior and not manager_unethical_response:
        return {
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
            "Робота із запереченнями": 10,
        }

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
    if features.get("client_driving_or_no_phone"):
        lines.append("Спроба презентації: клієнт не міг повноцінно взаємодіяти з телефоном, тому критерій зараховано за винятком.")
    elif presentation_level == "full":
        lines.append("Спроба презентації: менеджер назвав продукт або активність і пояснив суть чи де знайти інформацію, тому презентацію зараховано повністю.")
    elif presentation_level == "partial":
        lines.append("Спроба презентації: менеджер згадав продукт, активність або програму лояльності, але без повного розкриття суті.")
    else:
        lines.append("Спроба презентації: презентації продукту не було; інформація лише про бонус не рахується як презентація.")

    followup_type = features.get("followup_type", "none")
    if features.get("client_hung_up_interrupted"):
        lines.append("Домовленість про наступний контакт: клієнт завершив дзвінок завчасно, тому критерій зараховано за винятком.")
    elif int(features.get("followup_attempts_count") or 0) >= 2:
        lines.append("Домовленість про наступний контакт: менеджер зробив щонайменше дві окремі спроби домовитися про контакт, але клієнт не дав конкретики.")
    elif followup_type == "exact_time":
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
        elif call.get("call_completion_status") == "🟢 (слухавку поклав клієнт)" and features.get("client_hung_up_interrupted"):
            lines.append("Передзвон клієнту: розмова обірвалась з боку клієнта, тому окремий штраф за передзвон не застосовується.")
        else:
            lines.append("Передзвон клієнту: штрафу немає, додатковий передзвон у цьому сценарії не був потрібний.")
    elif scores.get("Передзвон клієнту", 0) == 10:
        lines.append("Передзвон клієнту: передзвон був, але не одразу, а протягом двох годин.")
    else:
        lines.append("Передзвон клієнту: потрібного передзвону не було, тому критерій не виконано.")

    if features.get("assumption_made"):
        lines.append("Не додумувати: менеджер припускав стан або намір клієнта без прямого підтвердження, тому критерій провалено.")
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

    card_elements = sum([
        bool(features.get("card_has_reason")),
        bool(features.get("card_has_followup_time")),
    ])
    if card_elements == 2:
        lines.append("Оформлення картки: у коментарі є причина незавершеної розмови та час наступного контакту.")
    elif card_elements == 1:
        lines.append("Оформлення картки: у коментарі є лише один з обов'язкових елементів: причина або час наступного контакту.")
    else:
        lines.append("Оформлення картки: у коментарі немає ані причини незавершеної розмови, ані часу наступного контакту.")

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
            lines.append("Утримання клієнта: реакція менеджера була формальною, без реальної спроби втримати клієнта.")
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
    interrupted_client_hangup = bool(features.get("client_hung_up_interrupted"))

    if logical_completion and has_farewell:
        if has_followup and not has_any_repeat:
            scores["Передзвон клієнту"] = 0
        return scores

    if status == "🟢 (слухавку поклав клієнт)":
        scores["Завершення розмови"] = 5
        if interrupted_client_hangup:
            scores["Домовленість про наступний контакт"] = 5
            if not has_any_repeat:
                scores["Передзвон клієнту"] = 0

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

    if has_followup and not has_any_repeat:
        scores["Передзвон клієнту"] = 0

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
        replacements, kb_data, kb_context = get_reference_data()
        
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

            analysis_result = analyze_call_cached(
                "openai" if run_openai else "claude",
                call["url"],
                call["call_date"],
                transcript,
                call["manager_comment"],
                kb_context,
                replacements,
                ANALYSIS_CACHE_VERSION,
            )

            if not analysis_result:
                st.warning("Помилка аналізу")
                continue

            clean_dialogue = analysis_result.get("cleaned_transcript") or transcript
            clean_dialogue = apply_replacements(clean_dialogue, replacements)
            features = analysis_result.get("features", {})
            features = normalize_presentation_level(features, clean_dialogue, kb_data)
            features = validate_bonus_features(features, clean_dialogue)
            features = validate_objection_and_retention(features, clean_dialogue)
            features = validate_professionalism_features(features, clean_dialogue)
            features = validate_forbidden_words(features, clean_dialogue)
            features = validate_assumption_made(features, clean_dialogue)
            comment = analysis_result.get("qa_comment", "").strip()

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
                if not call["ret_sheet_id"]:
                    st.error("Не обрано проєкт або менеджера РЕТ")
                    continue

                total_score = sum(scores.values())
                sheet_settings = get_manager_sheet_settings(call)

                try:
                    workbook = google_client.open_by_key(call["ret_sheet_id"])
                    scores_sheet = (
                        workbook.worksheet(sheet_settings["worksheet_name"])
                        if sheet_settings["worksheet_name"]
                        else workbook.sheet1
                    )
                except Exception as e:
                    st.error(f"Google error [manager workbook]: {e}")
                    continue

                try:
                    res = write_to_google_sheet(
                        scores_sheet,
                        call,
                        scores,
                        start_column=sheet_settings["start_column"],
                        start_row=sheet_settings["scores_start_row"],
                        criteria_start_row=sheet_settings["criteria_start_row"],
                    )
                    st.write("WRITE RESULT:", res)
                    if res is not True:
                        st.error(f"Google error [scores write]: {res}")
                except Exception as e:
                    st.error(f"Google error [scores write]: {e}")

                try:
                    append_manager_log(
                        scores_sheet,
                        call,
                        comment_for_sheet,
                        total_score,
                        ai_label,
                        start_row=sheet_settings["log_start_row"],
                    )
                except Exception as e:
                    st.error(f"Google error [manager log]: {e}")

                try:
                    log_workbook = google_client.open_by_key(LOG_SHEET_ID)
                except Exception as e:
                    st.error(f"Google error [QA logs workbook]: {e}")
                    log_workbook = None

                try:
                    if log_workbook is None:
                        raise RuntimeError("Не вдалося відкрити QA_LOGS_CALLS")
                    log_sheet = log_workbook.worksheet("Лист 1")
                    append_qa_log(
                        log_sheet,
                        call,
                        transcript,
                        clean_dialogue,
                        comment,
                        total_score
                    )
                except Exception as e:
                    st.error(f"Google error [QA log]: {e}")

                try:
                    if log_workbook is None:
                        raise RuntimeError("Не вдалося відкрити QA_LOGS_CALLS")
                    log_info_sheet = log_workbook.worksheet("LOG_INFO")
                    append_log_info(
                        log_info_sheet,
                        call,
                    )
                except Exception as e:
                    st.error(f"Google error [LOG_INFO]: {e}")

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

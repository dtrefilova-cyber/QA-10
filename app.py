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
    <h2 style="margin:0;">рџЋ§ QA-10</h2>
    <span style="color:#aaa;">РђРЅР°Р»С–Р· РґР·РІС–РЅРєС–РІ</span>
</div>
""", unsafe_allow_html=True)

check_date = st.date_input("Р”Р°С‚Р° РїРµСЂРµРІС–СЂРєРё", datetime.today())

qa_managers_list = [
    "Р”Р°СЂ'СЏ", "РќР°РґСЏ", "РќР°СЃС‚СЏ", "Р’Р»Р°РґРёРјРёСЂР°", "Р”С–Р°РЅР°", "Р СѓСЃР»Р°РЅР°", "РћР»РµРєСЃС–Р№"
]

call_completion_statuses = [
    "вљЄ (РІС–РґСЃСѓС‚РЅС–Р№ СЃС‚Р°С‚СѓСЃ)",
    "рџџў (СЃР»СѓС…Р°РІРєСѓ РїРѕРєР»Р°РІ РєР»С–С”РЅС‚)",
    "рџџЎ (С‚РµС…РЅС–С‡РЅС– РїСЂРѕР±Р»РµРјРё, Р·РІ'СЏР·РѕРє РѕР±С–СЂРІР°РІСЃСЏ)",
    "рџ”ґ (СЃР»СѓС…Р°РІРєСѓ РїРѕРєР»Р°РІ РјРµРЅРµРґР¶РµСЂ)",
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
    st.error(f"РџРѕРјРёР»РєР° Р·Р°РІР°РЅС‚Р°Р¶РµРЅРЅСЏ РјРµРЅРµРґР¶РµСЂС–РІ: {e}")

projects_list = sorted({item["project"] for item in managers_config})

if not managers_config:
    st.warning(
        "РЎРїРёСЃРѕРє РїСЂРѕС”РєС‚С–РІ С– РјРµРЅРµРґР¶РµСЂС–РІ РЅРµ Р·Р°РІР°РЅС‚Р°Р¶РёРІСЃСЏ Р· Р°СЂРєСѓС€Р° MANAGERS. "
        "РџРµСЂРµРІС–СЂС‚Рµ, С‰Рѕ РІ Р°СЂРєСѓС€С– С” Р·Р°РіРѕР»РѕРІРєРё MANAGERS_NAME, PROJECT, SHEET_ID "
        "С– С‰Рѕ РІ РєРѕР»РѕРЅС†С– SHEET_ID Р·Р°РїРѕРІРЅРµРЅС– Р·РЅР°С‡РµРЅРЅСЏ."
    )
    st.caption(
        f"Р”С–Р°РіРЅРѕСЃС‚РёРєР°: headers={managers_meta['headers']}, "
        f"header_row={managers_meta['header_row_index']}, "
        f"raw_rows={managers_meta['raw_rows_count']}, "
        f"valid_rows={managers_meta['valid_rows_count']}"
    )

# ================= INPUT =================
calls = []
for row in range(5):
    col1, col2 = st.columns(2)
    for col, idx in zip([col1, col2], [row * 2 + 1, row * 2 + 2]):
        with col.expander(f"рџ“ћ Р”Р·РІС–РЅРѕРє {idx}"):
            audio_url = st.text_input("РџРѕСЃРёР»Р°РЅРЅСЏ", key=f"url_{idx}")
            qa_manager = st.selectbox("QA", qa_managers_list, key=f"qa_{idx}")
            selected_project = st.selectbox(
                "РџСЂРѕС”РєС‚",
                projects_list,
                index=None,
                placeholder="РћР±РµСЂС–С‚СЊ РїСЂРѕС”РєС‚",
                key=f"project_{idx}",
                disabled=not projects_list
            )
            project_managers = [
                item for item in managers_config
                if item["project"] == selected_project
            ]
            manager_names = [item["manager_name"] for item in project_managers]
            selected_manager = st.selectbox(
                "РњРµРЅРµРґР¶РµСЂ Р Р•Рў",
                manager_names,
                index=None,
                placeholder="РћР±РµСЂС–С‚СЊ РјРµРЅРµРґР¶РµСЂР°",
                key=f"ret_{idx}",
                disabled=not manager_names
            )
            selected_manager_data = next(
                (item for item in project_managers if item["manager_name"] == selected_manager),
                None
            )
            client_id = st.text_input("ID", key=f"client_{idx}")
            call_date = st.text_input("Р”Р°С‚Р°", key=f"date_{idx}")
            bonus_check = st.selectbox(
                "Р‘РѕРЅСѓСЃ",
                ["РїСЂР°РІРёР»СЊРЅРѕ РЅР°СЂР°С…РѕРІР°РЅРѕ", "РїРѕРјРёР»РєРѕРІРѕ РЅР°СЂР°С…РѕРІР°РЅРѕ", "РЅРµ РїРѕС‚СЂС–Р±РЅРѕ"],
                key=f"bonus_{idx}"
            )
            repeat_col, completion_col = st.columns(2)
            with repeat_col:
                repeat_call = st.selectbox(
                    "РџРµСЂРµРґР·РІРѕРЅ",
                    ["С‚Р°Рє, Р±СѓРІ РїСЂРѕС‚СЏРіРѕРј РіРѕРґРёРЅРё", "С‚Р°Рє, Р±СѓРІ РїСЂРѕС‚СЏРіРѕРј 2 РіРѕРґРёРЅ", "РЅС–, РЅРµ Р±СѓР»Рѕ"],
                    key=f"repeat_{idx}"
                )
            with completion_col:
                call_completion_status = st.selectbox(
                    "Завершення виклику",
                    call_completion_statuses,
                    key=f"call_completion_{idx}"
                )
            manager_comment = st.text_area("РљРѕРјРµРЅС‚Р°СЂ", key=f"comment_{idx}")

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
            return {"ok": False, "error": "РќРµРјР°С” С‚СЂР°РЅСЃРєСЂРёРїС†С–С—", "transcript": None}

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

        parts = [f"РџСЂРѕРґСѓРєС‚: {name}"]
        if aliases:
            parts.append(f"РђР»С–Р°СЃРё: {aliases}")
        if description:
            parts.append(f"РћРїРёСЃ: {description}")

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
        "Р·Р°Р»РёС€С‚Рµ РїРѕРІС–РґРѕРјР»РµРЅРЅСЏ",
        "РїС–СЃР»СЏ СЃРёРіРЅР°Р»Сѓ",
        "Р°Р±РѕРЅРµРЅС‚ РЅРµРґРѕСЃС‚СѓРїРЅРёР№",
        "РЅРµ РјРѕР¶Рµ РІС–РґРїРѕРІС–СЃС‚Рё",
        "voice mail",
        "voicemail",
        "please leave a message",
        "РЅРѕРјРµСЂ РЅРµ РѕР±СЃР»СѓРіРѕРІСѓС”С‚СЊСЃСЏ"
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

        "assumption_made": False,

        "comment_match_level": "none",
        "comment_complete": False
    }

    for k, v in defaults.items():
        features.setdefault(k, v)

    return features


def build_dictionary_context(replacements):
    if not replacements:
        return "РЎР»РѕРІРЅРёРє Р·Р°РјС–РЅ РЅРµ РїРµСЂРµРґР°РЅРёР№."

    return "\n".join([f"{k} в†’ {v}" for k, v in replacements.items()])


def get_analysis_output_schema():
    return """
РџРѕРІРµСЂРЅРё ONLY valid JSON С‚Р°РєРѕРіРѕ С„РѕСЂРјР°С‚Сѓ:
{
  "cleaned_transcript": "РѕС‡РёС‰РµРЅРёР№ РґС–Р°Р»РѕРі",
  "qa_comment": "РіРѕС‚РѕРІРёР№ QA-РєРѕРјРµРЅС‚Р°СЂ РїРѕ РєСЂРёС‚РµСЂС–СЏС…, РєРѕР¶РµРЅ РєСЂРёС‚РµСЂС–Р№ Р· РЅРѕРІРѕРіРѕ СЂСЏРґРєР°",
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
    "speech_quality": "bad" | "good"
  }
}
"""


def build_combined_analysis_prompt(prompt_body, raw_dialogue, replacements):
    dictionary_context = build_dictionary_context(replacements)
    return f"""
{prompt_body}

---------------------
РЎР›РћР’РќРРљ Р—РђРњР†Рќ
---------------------

РЎР»РѕРІРЅРёРє Р·Р°РјС–РЅ С” РћР‘РћР’'РЇР—РљРћР’РРњ.
РЇРєС‰Рѕ СЃР»РѕРІРѕ Р°Р±Рѕ С„СЂР°Р·Р° С” Сѓ СЃР»РѕРІРЅРёРєСѓ, РІРёРєРѕСЂРёСЃС‚РѕРІСѓР№ С‚С–Р»СЊРєРё РІР°СЂС–Р°РЅС‚ Р·С– СЃР»РѕРІРЅРёРєР°.
РќРµ РІРёРіР°РґСѓР№ РІР»Р°СЃРЅРёС… РІР°СЂС–Р°РЅС‚С–РІ, СЏРєС‰Рѕ СЃР»РѕРІРѕ С” Сѓ СЃР»РѕРІРЅРёРєСѓ.

{dictionary_context}

---------------------
РћР§РРЎРўРљРђ РўР РђРќРЎРљР РРџРўРЈ
---------------------

РЎРїРѕС‡Р°С‚РєСѓ РѕС‡РёСЃС‚Рё С‚СЂР°РЅСЃРєСЂРёРїС‚:
- РІРёРїСЂР°РІ РїРѕРјРёР»РєРё СЂРѕР·РїС–Р·РЅР°РІР°РЅРЅСЏ
- Р·Р°СЃС‚РѕСЃСѓР№ СЃР»РѕРІРЅРёРє Р·Р°РјС–РЅ
- РЅРµ Р·РјС–РЅСЋР№ СЃРµРЅСЃ
- РЅРµ СЃРєРѕСЂРѕС‡СѓР№ С‚РµРєСЃС‚
- Р·Р°РјС–РЅРё ch_0 РЅР° "РњРµРЅРµРґР¶РµСЂ", ch_1 РЅР° "РљР»С–С”РЅС‚"

РџС–СЃР»СЏ С†СЊРѕРіРѕ:
- РїСЂРѕР°РЅР°Р»С–Р·СѓР№ РІР¶Рµ РѕС‡РёС‰РµРЅРёР№ С‚СЂР°РЅСЃРєСЂРёРїС‚
- СЃС„РѕСЂРјСѓР№ РіРѕС‚РѕРІРёР№ qa_comment Сѓ С‚РѕРјСѓ Р¶ Р·Р°РїРёС‚С–
- qa_comment РјР°С” Р±СѓС‚Рё СѓРєСЂР°С—РЅСЃСЊРєРѕСЋ, РїРѕ РѕРґРЅРѕРјСѓ РєСЂРёС‚РµСЂС–СЋ РЅР° СЂСЏРґРѕРє

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

    # СЏРєС‰Рѕ Р°РІС‚РѕРІС–РґРїРѕРІС–РґР°С‡ в†’ РІСЃС– 0
    if dialogue and is_autoresponder(dialogue):
        return {
            "Р’СЃС‚Р°РЅРѕРІР»РµРЅРЅСЏ РєРѕРЅС‚Р°РєС‚Сѓ": 0,
            "РЎРїСЂРѕР±Р° РїСЂРµР·РµРЅС‚Р°С†С–С—": 0,
            "Р”РѕРјРѕРІР»РµРЅС–СЃС‚СЊ РїСЂРѕ РЅР°СЃС‚СѓРїРЅРёР№ РєРѕРЅС‚Р°РєС‚": 0,
            "РџСЂРѕРїРѕР·РёС†С–СЏ Р±РѕРЅСѓСЃСѓ": 0,
            "Р—Р°РІРµСЂС€РµРЅРЅСЏ СЂРѕР·РјРѕРІРё": 0,
            "РџРµСЂРµРґР·РІРѕРЅ РєР»С–С”РЅС‚Сѓ": 0,
            "РќРµ РґРѕРґСѓРјСѓРІР°С‚Рё": 0,
            "РЇРєС–СЃС‚СЊ РјРѕРІР»РµРЅРЅСЏ": 0,
            "РџСЂРѕС„РµСЃС–РѕРЅР°Р»С–Р·Рј": 0,
            "РћС„РѕСЂРјР»РµРЅРЅСЏ РєР°СЂС‚РєРё": 0,
            "РЈС‚СЂРёРјР°РЅРЅСЏ РєР»С–С”РЅС‚Р°": 0,
            "Р РѕР±РѕС‚Р° С–Р· Р·Р°РїРµСЂРµС‡РµРЅРЅСЏРјРё": 0
        }

    # ---------------- РљРѕРЅС‚Р°РєС‚ ----------------
    elements = sum([
    f["manager_name_present"],
    f["manager_position_present"],
    f["company_present"],
    f["client_name_used"],
    f["purpose_present"],
    f.get("friendly_question", False)
])

    s["Р’СЃС‚Р°РЅРѕРІР»РµРЅРЅСЏ РєРѕРЅС‚Р°РєС‚Сѓ"] = (
        7.5 if elements >= 4 else
        5 if elements == 3 else
        2.5 if elements == 2 else
        0
    )

    # ---------------- РЎРїСЂРѕР±Р° РїСЂРµР·РµРЅС‚Р°С†С–С— ----------------
    level = f.get("presentation_level", "none")

    if level == "full":
        s["РЎРїСЂРѕР±Р° РїСЂРµР·РµРЅС‚Р°С†С–С—"] = 5
    elif level == "partial":
        s["РЎРїСЂРѕР±Р° РїСЂРµР·РµРЅС‚Р°С†С–С—"] = 2.5
    else:
        s["РЎРїСЂРѕР±Р° РїСЂРµР·РµРЅС‚Р°С†С–С—"] = 0

    # ---------------- Р”РѕРјРѕРІР»РµРЅС–СЃС‚СЊ ----------------
    fup = f.get("followup_type", "none")
    s["Р”РѕРјРѕРІР»РµРЅС–СЃС‚СЊ РїСЂРѕ РЅР°СЃС‚СѓРїРЅРёР№ РєРѕРЅС‚Р°РєС‚"] = (
        5 if fup == "exact_time"
        else 2.5 if fup == "offer"
        else 0
    )

    # ---------------- Р‘РѕРЅСѓСЃ ----------------
    if not f.get("bonus_offered"):
        s["РџСЂРѕРїРѕР·РёС†С–СЏ Р±РѕРЅСѓСЃСѓ"] = 0
    else:
        bonus_conditions = sum([
            bool(f.get("bonus_has_type")),
            bool(f.get("bonus_has_duration")),
            bool(f.get("bonus_has_value"))
        ])
        s["РџСЂРѕРїРѕР·РёС†С–СЏ Р±РѕРЅСѓСЃСѓ"] = 10 if bonus_conditions >= 2 else 5

    # ---------------- Р—Р°РІРµСЂС€РµРЅРЅСЏ ----------------
    s["Р—Р°РІРµСЂС€РµРЅРЅСЏ СЂРѕР·РјРѕРІРё"] = 5 if f.get("has_farewell") else 0

    # ---------------- РџРµСЂРµРґР·РІРѕРЅ ----------------
    repeat = meta["repeat_call"]
    
    if fup in ["none", "offer", "exact_time"]:
        s["РџРµСЂРµРґР·РІРѕРЅ РєР»С–С”РЅС‚Сѓ"] = 15
    else:
        s["РџРµСЂРµРґР·РІРѕРЅ РєР»С–С”РЅС‚Сѓ"] = (
            15 if repeat == "С‚Р°Рє, Р±СѓРІ РїСЂРѕС‚СЏРіРѕРј РіРѕРґРёРЅРё"
            else 10 if repeat == "С‚Р°Рє, Р±СѓРІ РїСЂРѕС‚СЏРіРѕРј 2 РіРѕРґРёРЅ"
            else 0
        )

    # ---------------- РќРµ РґРѕРґСѓРјСѓРІР°С‚Рё ----------------
    if f.get("assumption_made"):
        s["РќРµ РґРѕРґСѓРјСѓРІР°С‚Рё"] = 2.5
    else:
        s["РќРµ РґРѕРґСѓРјСѓРІР°С‚Рё"] = 5

    # ---------------- РЇРєС–СЃС‚СЊ РјРѕРІР»РµРЅРЅСЏ ----------------
    quality = f.get("speech_quality", "bad")

    if quality == "good":
        s["РЇРєС–СЃС‚СЊ РјРѕРІР»РµРЅРЅСЏ"] = 2.5
    else:
        s["РЇРєС–СЃС‚СЊ РјРѕРІР»РµРЅРЅСЏ"] = 0

    # ---------------- РџСЂРѕС„РµСЃС–РѕРЅР°Р»С–Р·Рј ----------------
    s["РџСЂРѕС„РµСЃС–РѕРЅР°Р»С–Р·Рј"] = (
        5 if meta["bonus_check"] == "РїРѕРјРёР»РєРѕРІРѕ РЅР°СЂР°С…РѕРІР°РЅРѕ" else 10
    )

    # ---------------- РљР°СЂС‚РєР° ----------------
    match = f.get("comment_match_level", "none")
    complete = f.get("comment_complete", False)

    if match == "none":
        s["РћС„РѕСЂРјР»РµРЅРЅСЏ РєР°СЂС‚РєРё"] = 0
    elif not complete:
        s["РћС„РѕСЂРјР»РµРЅРЅСЏ РєР°СЂС‚РєРё"] = 2.5
    else:
        s["РћС„РѕСЂРјР»РµРЅРЅСЏ РєР°СЂС‚РєРё"] = 5

    # ---------------- РЈС‚СЂРёРјР°РЅРЅСЏ ----------------
    lvl = f.get("continuation_level", "none")

    if not f.get("client_wants_to_end"):
        behavior = f.get("continuation_behavior", "neutral")
        s["РЈС‚СЂРёРјР°РЅРЅСЏ РєР»С–С”РЅС‚Р°"] = (
            20 if behavior == "active"
            else 15 if behavior == "neutral"
            else 10 if behavior == "passive"
            else 0
        )
    else:
        s["РЈС‚СЂРёРјР°РЅРЅСЏ РєР»С–С”РЅС‚Р°"] = (
            20 if lvl == "strong"
            else 15 if lvl == "weak"
            else 10 if lvl == "formal"
            else 5 if lvl == "none"
            else 0
        )

    # ---------------- Р—Р°РїРµСЂРµС‡РµРЅРЅСЏ ----------------
    if not f.get("objection_detected"):
        s["Р РѕР±РѕС‚Р° С–Р· Р·Р°РїРµСЂРµС‡РµРЅРЅСЏРјРё"] = 10
    else:
        s["Р РѕР±РѕС‚Р° С–Р· Р·Р°РїРµСЂРµС‡РµРЅРЅСЏРјРё"] = (
            10 if lvl == "strong"
            else 5 if lvl == "weak"
            else 0
        )

    return s


def format_comment_for_sheet(comment):
    if not comment:
        return ""

    lines = [line.strip() for line in str(comment).splitlines() if line.strip()]
    return " | ".join(lines)

# ================= RUN =================
if "results" not in st.session_state:
    st.session_state["results"] = []

col1, col2 = st.columns(2)
run_openai = col1.button("рџљЂ OpenAI", type="primary")
run_claude = col2.button("рџ§  Claude")

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

        with st.spinner(f"РђРЅР°Р»С–Р· РґР·РІС–РЅРєР° {i+1}..."):

            transcript = transcribe_audio(call["url"])
            if not transcript:
                st.warning("РќРµРјР°С” С‚СЂР°РЅСЃРєСЂРёРїС†С–С—")
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
                st.warning("РџРѕРјРёР»РєР° Р°РЅР°Р»С–Р·Сѓ")
                continue

            clean_dialogue = analysis_result.get("cleaned_transcript") or transcript
            clean_dialogue = apply_replacements(clean_dialogue, replacements)
            features = analysis_result.get("features", {})
            comment = analysis_result.get("qa_comment", "").strip()
            presentation_detected = detect_presentation(clean_dialogue, kb_data)

            # С„С–Р»СЊС‚СЂ С‡РµСЂРµР· Р±Р°Р·Сѓ Р·РЅР°РЅСЊ
            if not presentation_detected:
                features["presentation_level"] = "none"

            if not features:
                st.warning("РџРѕРјРёР»РєР° Р°РЅР°Р»С–Р·Сѓ")
                continue

            scores = score_call(features, call, clean_dialogue)
            if not comment:
                comment = "РџРѕРјРёР»РєР° РіРµРЅРµСЂР°С†С–С— РєРѕРјРµРЅС‚Р°СЂСЏ"
            comment_for_sheet = format_comment_for_sheet(comment)
            ai_label = "OpenAI" if run_openai else "Claude"

            st.session_state["results"].append({
                "scores": scores,
                "comment": comment
            })

            if google_client:
                try:
                    if not call["ret_sheet_id"]:
                        st.error("РќРµ РѕР±СЂР°РЅРѕ РїСЂРѕС”РєС‚ Р°Р±Рѕ РјРµРЅРµРґР¶РµСЂР° Р Р•Рў")
                        continue

                    # рџџў С‚Р°Р±Р»РёС†СЏ РјРµРЅРµРґР¶РµСЂР°
                    sheet = google_client.open_by_key(call["ret_sheet_id"]).sheet1

                    # рџџў С„РѕСЂРјСѓС”РјРѕ РѕС†С–РЅРєСѓ РѕРґРЅРёРј СЂСЏРґРєРѕРј
                    total_score = sum(scores.values())

                    # рџџў СЃРїРѕС‡Р°С‚РєСѓ РѕС†С–РЅРєРё
                    res = write_to_google_sheet(sheet, call, scores) 
                    st.write("WRITE RESULT:", res)

                    # рџџў Р·Р°РїРёСЃ Сѓ С‚Р°Р±Р»РёС†СЋ РјРµРЅРµРґР¶РµСЂР° (С‚РІРѕСЏ СЃС‚СЂСѓРєС‚СѓСЂР°)
                    append_manager_log(
                        sheet,
                        call,
                        comment_for_sheet,
                        total_score,
                        ai_label
                    )

                    # рџџў Р»РѕРі С‚Р°Р±Р»РёС†СЏ
                    log_sheet = google_client.open_by_key(LOG_SHEET_ID).worksheet("Р›РёСЃС‚ 1")
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
    with st.expander(f"рџ“ћ Р”Р·РІС–РЅРѕРє {i+1}", expanded=(i == 0)):
        df = pd.DataFrame(
            list(res["scores"].items()),
            columns=["РљСЂРёС‚РµСЂС–Р№", "РћС†С–РЅРєР°"]
        )
        df["РћС†С–РЅРєР°"] = df["РћС†С–РЅРєР°"].apply(lambda x: f"{float(x):.1f}")
        st.table(df)

        total = sum(res["scores"].values())
        st.success(f"Р—Р°РіР°Р»СЊРЅРёР№ Р±Р°Р»: {total:.1f}")

        st.markdown("### рџ’¬ РљРѕРјРµРЅС‚Р°СЂ QA")
        for line in res["comment"].split("\n"):     
            st.write(line)

# ================= EXPORT =================
if st.session_state["results"]:
    xls = BytesIO()
    with pd.ExcelWriter(xls, engine="openpyxl") as writer:
        for i, res in enumerate(st.session_state["results"]):
            df = pd.DataFrame(res["scores"].items(), columns=["РљСЂРёС‚РµСЂС–Р№", "РћС†С–РЅРєР°"])
            df.to_excel(writer, sheet_name=f"Call_{i+1}", index=False)
    xls.seek(0)

    st.download_button(
        label="рџ“Ґ Р—Р°РІР°РЅС‚Р°Р¶РёС‚Рё Excel",
        data=xls,
        file_name="qa_results.xlsx"
    )

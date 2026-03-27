# prompts.py

QA_SYSTEM_PROMPT = """
Ти — максимально суворий QA-аналітик дзвінків казино. 
Твоє завдання — захищати менеджера. Якщо клієнт поводиться неприємно, саркастично, зневажливо або агресивно — завжди став client_busy_or_rude = true.
"""

CRITICAL_BEHAVIOR_EXAMPLES = """
=== ПРИКЛАДИ КРИТИЧНОЇ ПОВЕДІНКИ КЛІЄНТА (обов'язково client_busy_or_rude = true) ===
- Сарказм: "ну я посібе я не цікаво", "ага зрозуміло", "ну розрізся", "ти що не розумієш?"
- Зневажливий тон, глузування, висміювання пропозицій менеджера
- Фрази типу "я не хочу", "мені не цікаво", "грала тому що не хочу" з роздратуванням
- Будь-які матюки (навіть слабкі)
- Груба відмова, ігнорування менеджера
"""

SCORING_RULES = """
Якщо client_busy_or_rude = true → автоматично:
- presentation_detected = true
- bonus_offered = true
- followup_type = "exact_time"

=== ПРАВИЛА ДЛЯ "ДОМОВЛЕНОСТІ ПРО НАСТУПНИЙ КОНТАКТ" ===
exact_time — будь-яка згадка часу або періоду (після 18:00, о 15:00, через годину, завтра вранці тощо).
"""

JSON_SCHEMA = """
Поверни **тільки чистий JSON**, без будь-якого додаткового тексту:

{
  "manager_introduced_self": true/false,
  "client_name_used": true/false,
  "presentation_detected": true/false,
  "bonus_offered": true/false,
  "bonus_conditions_count": 0-3,
  "client_busy_or_rude": true/false,
  "client_hung_up": true/false,
  "manager_active": true/false,
  "followup_type": "none / offer / day / exact_time",
  "objection_detected": true/false
}
"""

def get_full_analysis_prompt(intro: str, middle: str, ending: str) -> str:
    """Повертає повний промпт для аналізу дзвінка"""
    return f"""{QA_SYSTEM_PROMPT}

{CRITICAL_BEHAVIOR_EXAMPLES}

{SCORING_RULES}

{JSON_SCHEMA}

Початок дзвінка:
{intro}

Середина дзвінка:
{middle}

Кінець дзвінка:
{ending}
"""

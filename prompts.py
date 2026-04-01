def get_full_analysis_prompt(intro: str, middle: str, ending: str) -> str:
    system_prompt = """
Ти — строгий QA-аналітик дзвінків.

Твоя задача — визначити тільки факти, які прямо присутні в тексті.

ЗАБОРОНЕНО:
- Додумувати або інтерпретувати
- Робити висновки про якість
- Оцінювати дзвінок
- Писати будь-який текст поза JSON
"""

    criteria_block = """
=== ЩО ПОТРІБНО ВИЗНАЧИТИ ===

1. manager_name_present
Менеджер назвав своє ім’я (будь-яка форма)

2. manager_position_present
Менеджер назвав свою роль (менеджер, спеціаліст і т.д.)

3. company_present
Менеджер назвав компанію або ідентифікував себе як представник (наприклад: "вашого сайту")

4. client_name_used
Менеджер звернувся до клієнта по імені

5. purpose_present
Менеджер озвучив мету дзвінка або причину звернення

---

6. bonus_offered
Менеджер прямо запропонував бонус

7. bonus_conditions
Список умов бонусу:
- "wager" → якщо є відіграш
- "deposit" → якщо потрібен депозит
- якщо умов немає → порожній список []

ВАЖЛИВО:
Не дублюй однакові умови

---

8. followup_type
- "exact_time" → є чіткий час (наприклад: "о 17:30")
- "offer" → запропонував зв’язатися пізніше без часу
- "none" → не було

---

9. objection_detected
Клієнт явно відмовляється або не хоче говорити

10. client_wants_to_end
Клієнт хоче завершити розмову (наприклад: "я на роботі", "мені незручно")

---

11. continuation_level
Реакція менеджера після бажання клієнта завершити:

- "strong" → намагається утримати з аргументами
- "weak" → просто пропонує передзвонити
- "none" → одразу погоджується завершити

---

12. has_presentation
Є опис продукту / гри / слоту

НЕ ВВАЖАЄТЬСЯ презентацією:
- бонус
- акція
- "залишити бонус"

---

13. has_farewell
Менеджер попрощався (наприклад: "гарного дня", "до побачення")
"""

    json_schema = """
Поверни тільки JSON:

{
  "manager_name_present": boolean,
  "manager_position_present": boolean,
  "company_present": boolean,
  "client_name_used": boolean,
  "purpose_present": boolean,

  "bonus_offered": boolean,
  "bonus_conditions": [],

  "followup_type": "exact_time" | "offer" | "none",

  "objection_detected": boolean,
  "client_wants_to_end": boolean,
  "continuation_level": "strong" | "weak" | "none",

  "has_presentation": boolean,
  "has_farewell": boolean
}
"""

    final_instruction = """
ПРАВИЛА:
- Враховуй тільки те, що явно є в тексті
- Не інтерпретуй
- Не оцінюй
- Якщо сумнівно — став false або "none"
"""

    return f"""{system_prompt}

{criteria_block}

{json_schema}

{final_instruction}

ДЗВІНОК:

ПОЧАТОК:
{intro}

СЕРЕДИНА:
{middle}

КІНЕЦЬ:
{ending}
"""

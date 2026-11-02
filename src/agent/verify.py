"""Защита от выдумок: каждое число в ответе агента должно найтись в выводах его инструментов."""
import re

NUM = re.compile(r"(?<![\w.])[−-]?(\d{1,3}(?:[ \u00a0]\d{3})+(?:[.,]\d+)?|\d+(?:[.,]\d+)?)(?![\w])")
UNIT = re.compile(r"\s*(%|млн|₽|руб|п\.п|тыс)")


def _val(tok):
    return float(re.sub(r"[ \u00a0]", "", tok).replace(",", "."))


def tool_numbers(outputs):
    vals = set()
    for o in outputs:
        for m in re.finditer(r"-?\d+(?:\.\d+)?", o):
            vals.add(abs(float(m.group())))
    return vals


def unverified(answer, outputs):
    """Числа из ответа, которых нет в выводах инструментов (допуск на округление 0,05)."""
    known, bad = tool_numbers(outputs), []
    for m in NUM.finditer(answer):
        txt, v = m.group(1), _val(m.group(1))
        start = m.start()
        line_start = answer.rfind("\n", 0, start) + 1
        if answer[line_start:start].strip(" -*") == "" and answer[m.end():m.end() + 1] in (".", ")"):
            continue                                  # номер пункта списка
        if v == 2025 or (v == int(v) and v <= 12 and not UNIT.match(answer[m.end():m.end() + 8])):
            continue                                  # год, месяцы, небольшие счётчики без единицы
        if not any(abs(v - k) <= max(0.051, 0.0005 * k) for k in known):
            bad.append(txt)
    return sorted(set(bad), key=bad.index)

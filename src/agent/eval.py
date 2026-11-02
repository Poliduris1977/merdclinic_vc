"""
Проверка агента на вопросах с известным ответом (нужен ANTHROPIC_API_KEY): python src/agent/eval.py [--model ...]
Зачёт по вопросу: вызван хотя бы один из нужных инструментов, в ответе есть все ожидаемые признаки,
нет чисел, которых нет в выводах инструментов. Ключ ответов (answer_key) не читается: ожидания заданы здесь.
Результат: data/agent_eval/results.md (годится как основа сравнения моделей: тот же набор, другой --model).
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import agent  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
# (вопрос, нужные инструменты (любой из), признаки в ответе (все), что проверяем)
CASES = [
    ("Почему валовая маржа в мае 2025 ниже плана?", {"top_drivers", "margin_bridge"}, [r"(?i)стоматолог|DENT", r"(?i)объ[её]м"], "B1: провал объёма стоматологии"),
    ("Что произошло с лазерной коррекцией зрения в клинике C03 летом?", {"top_drivers"}, [r"C03|Юг", r"(?i)объ[её]м"], "B2: простой оборудования, объём"),
    ("Почему косметология не выполняет план по выручке с апреля?", {"top_drivers", "margin_bridge"}, [r"(?i)косметолог|COSM", r"(?i)цен|прайс|индексац"], "B3: индексация +8% против +4%"),
    ("Сколько на самом деле стоят скидки по сети за год?", {"margin_bridge"}, [r"(?i)скидк"], "B4: ценовой эффект скидок"),
    ("Почему деньги на счетах выросли сильнее прибыли в 2025 году?", {"cashflow"}, [r"(?i)амортизац", r"(?i)долг|кредитор"], "мост прибыль → деньги"),
    ("Какие ошибки нашла сверка 1С и ERP и что осталось нерешённым?", {"data_quality"}, [r"(?i)без врача", r"120"], "сверка: остаток 2 визита, 120 700 ₽"),
    ("Какой налог на прибыль мы заплатили в 2025 году?", set(), [r"(?i)нет данных|не (учит|включ|входит)|вне"], "границы данных: агент не должен выдумывать"),
    ("Сравни маржу после оплаты врачей по грейдам в сети", {"doctors"}, [r"(?i)lead|senior|middle|junior"], "аналитика по врачам"),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model")
    a = ap.parse_args()
    rows, passed = [], 0
    for q, need, signs, what in CASES:
        rec = agent.run(q, model=a.model)
        used = {c["tool"] for c in rec["tool_calls"]}
        t_ok = (not need) or bool(need & used)
        s_ok = all(re.search(p, rec["answer"]) for p in signs)
        n_ok = not rec["unverified_numbers"]
        ok = t_ok and s_ok and n_ok
        passed += ok
        rows.append((ok, q, what, sorted(used), t_ok, s_ok, rec["unverified_numbers"], rec["usage"], rec["answer"]))
        print(("OK   " if ok else "FAIL ") + what)
    out = ROOT / "data" / "agent_eval"
    out.mkdir(parents=True, exist_ok=True)
    md = [f"# Проверка агента: {passed} из {len(CASES)}", f"Модель: {a.model or agent.DEFAULT_MODEL}", "",
          "| Зачёт | Что проверяем | Инструменты | Признаки | Непроверенные числа | Токены вх/вых |", "|---|---|---|---|---|---|"]
    for ok, q, what, used, t_ok, s_ok, bad, u, ans in rows:
        md.append(f"| {'да' if ok else 'нет'} | {what} | {', '.join(used) or '-'}{'' if t_ok else ' (нет нужного)'} | {'да' if s_ok else 'нет'} | {', '.join(bad) or '-'} | {u['input']}/{u['output']} |")
    md += ["", "## Ответы", ""]
    for ok, q, what, *_, ans in rows:
        md += [f"### {q}", "", ans, ""]
    (out / "results.md").write_text("\n".join(md), encoding="utf-8")
    print(f"\n{passed} из {len(CASES)}. Подробности: {out / 'results.md'}")


if __name__ == "__main__":
    main()

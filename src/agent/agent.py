"""
ИИ-агент отчётности клиник на Claude (tool use). Модель выбирает инструменты и объясняет числа, считает код.

Запуск:  export ANTHROPIC_API_KEY=...
         python src/agent/agent.py "Почему упала валовая маржа в мае?"
         python src/agent/agent.py --summary "Резюме для директора за 2025 год"
Модель: переменная CLINIC_AGENT_MODEL или --model (по умолчанию claude-sonnet-5-5).
Журнал прогонов: logs/agent_runs.jsonl (вопрос, вызовы инструментов, токены, ответ, непроверенные числа).
"""
import argparse
import json
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import tools  # noqa: E402
from verify import unverified  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
LOG = ROOT / "logs" / "agent_runs.jsonl"
DEFAULT_MODEL = "claude-sonnet-5-5"
MAX_STEPS = 8

SYSTEM = """Ты аналитик управленческой отчётности сети из трёх клиник (C01 «Невский», C02 «Север», C03 «Юг»), направления: \
DENT стоматология, FAM семейная медицина, COSM косметология, LASER лазерная коррекция зрения, PLAS пластическая хирургия. \
Данные за 2025 год: план и факт, суммы в млн ₽. Данные синтетические (демо).

Правила:
1. Все числа берёшь только из результатов инструментов. Сам ничего не считаешь, не суммируешь и не округляешь сильнее, чем в выводе. Если нужного числа нет, вызови инструмент ещё раз или скажи, что данных нет.
2. В ответе всегда указывай период и клинику, к которым относится число.
3. «Валовая маржа» (после оплаты врачей, персонала, амортизации) и «прямая маржа» (до оплаты труда, её даёт top_drivers) разные вещи, не смешивай.
4. На вопрос «почему упало» вызывай pnl (для вопросов про худший или лучший месяц и динамику: by_month=true), затем top_drivers (сначала по направлениям, при необходимости по direction_clinic), margin_bridge и events. Отделяй факт (что показали данные) от гипотезы. Причину называй фактом только если событие есть в events и совпадает по периоду и направлению; всё остальное (простой, индексация цен и т.п.) это гипотеза, предложи подтвердить у управляющих. Эффекты top_drivers считаются до оплаты труда (прямая маржа) и не равны отклонению валовой маржи: не сравнивай их как одно число. Формулировки «худший», «почти весь», «главный» используй только если это следует из чисел инструментов (доля считается инструментом, а не тобой).
5. Отклонения в мосте оплаты врачей («Врачи: объём» и «цена и скидки») положительны из-за того, что оплата врачей идёт от выручки и снижается вместе с ней: это следствие, а не экономия.
6. Отвечай по-русски, коротко: сначала вывод одной-двумя фразами, затем 2-4 факта с числами. Без вводных и без советов, о которых не просили.
7. Расчёты ДМС, налогов на прибыль и НДС вне данных. Если спрашивают, скажи прямо."""

SUMMARY_MODE = """Режим: резюме для директора. Не больше 150 слов. Структура: итог одной строкой; 3 главные причины с числами; 1 риск; 1 предложение, что проверить. \
Собери факты инструментами и сохрани текст через write_report (имя файла по теме)."""


def run(question, summary=False, client=None, model=None, quiet=True, log=True):
    if client is None:
        import anthropic
        client = anthropic.Anthropic()
    model = model or os.environ.get("CLINIC_AGENT_MODEL", DEFAULT_MODEL)
    system = SYSTEM + ("\n\n" + SUMMARY_MODE if summary else "")
    messages = [{"role": "user", "content": question}]
    calls, outputs, usage, t0 = [], [], {"input": 0, "output": 0}, time.time()
    answer = ""
    for _ in range(MAX_STEPS):
        resp = client.messages.create(model=model, max_tokens=2000, system=system, tools=tools.TOOLS, messages=messages)
        usage["input"] += resp.usage.input_tokens
        usage["output"] += resp.usage.output_tokens
        messages.append({"role": "assistant", "content": resp.content})
        uses = [b for b in resp.content if b.type == "tool_use"]
        answer = "".join(b.text for b in resp.content if b.type == "text") or answer
        if resp.stop_reason != "tool_use" or not uses:
            break
        results = []
        for u in uses:
            out = tools.call(u.name, u.input)
            calls.append({"tool": u.name, "args": u.input})
            if u.name != "write_report":
                outputs.append(out)
            if not quiet:
                print(f"  [{u.name}] {json.dumps(u.input, ensure_ascii=False)}", file=sys.stderr)
            results.append({"type": "tool_result", "tool_use_id": u.id, "content": out})
        messages.append({"role": "user", "content": results})
    else:
        answer += "\n\n(Остановлено: превышен лимит шагов.)"
    bad = unverified(answer, outputs)
    rec = {"ts": time.strftime("%Y-%m-%d %H:%M:%S"), "model": model, "question": question, "summary_mode": summary, "tool_calls": calls,
           "usage": usage, "seconds": round(time.time() - t0, 1), "answer": answer, "unverified_numbers": bad}
    if log:
        LOG.parent.mkdir(exist_ok=True)
        with LOG.open("a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    return rec


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("question")
    ap.add_argument("--summary", action="store_true", help="резюме для директора (сохраняется в data/agent_reports)")
    ap.add_argument("--model")
    a = ap.parse_args()
    rec = run(a.question, a.summary, model=a.model, quiet=False)
    print(rec["answer"])
    if rec["unverified_numbers"]:
        print(f"\n⚠ Числа, которых нет в выводах инструментов: {', '.join(rec['unverified_numbers'])}. Проверьте вручную.")
    print(f"\n[{rec['model']}: {len(rec['tool_calls'])} вызовов, токены {rec['usage']['input']}/{rec['usage']['output']}, {rec['seconds']} с]")


if __name__ == "__main__":
    main()

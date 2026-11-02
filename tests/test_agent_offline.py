"""Офлайн-тесты агента: инструменты против известных чисел, проверка выдумок, цикл агента на заглушке модели. Запуск: python tests/test_agent_offline.py"""
import json
import sys
from pathlib import Path
from types import SimpleNamespace as NS

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src" / "agent"))
import agent  # noqa: E402
import tools  # noqa: E402
from verify import unverified  # noqa: E402

ok = []


def check(name, cond):
    ok.append(cond)
    print(("OK   " if cond else "FAIL ") + name)


J = lambda n, a: json.loads(tools.call(n, a))  # noqa: E731

# 1. инструменты против контрольных чисел Excel
p = J("pnl", {"period": "2025"})
check("pnl 2025: валовая маржа 291,43 -> 268,10", (p["валовая_маржа"]["план"], p["валовая_маржа"]["факт"]) == (291.43, 268.1))
check("pnl 2025: прибыль 179,56 -> 158,05", (p["прибыль_до_налога"]["план"], p["прибыль_до_налога"]["факт"]) == (179.56, 158.05))
b = J("margin_bridge", {"period": "2025"})
check("мост: сумма факторов = отклонение валовой маржи", abs(sum(f["эффект"] for f in b["факторы"]) - b["валовая_маржа"]["отклонение"]) <= 0.05)
cf = J("cashflow", {"period": "2025"})
check("ДДС: +185,04, остаток 216,04", (cf["чистое_изменение"], cf["остаток_на_конец"]) == (185.04, 216.04))
check("мост денег: сумма = чистое изменение", abs(sum(x["эффект"] for x in cf["мост_прибыль_в_деньги"]) - cf["чистое_изменение"]) <= 0.05)
sc = sum(J("pnl", {"period": "2025", "clinic": c})["выручка"]["факт"] for c in ("C01", "C02", "C03"))
check("выручка клиник = сеть", abs(sc - p["выручка"]["факт"]) <= 0.03)
may = J("top_drivers", {"period": "2025-05", "by": "direction"})["худшие"][0]
check("май: первое по просадке направление DENT, объём < −10%", may["срез"] == "DENT" and may["объём_к_плану_%"] < -10)
lz = [x for x in J("top_drivers", {"period": "2025-07:2025-09", "by": "direction_clinic", "n": 5})["худшие"] if x["срез"] == "LASER / C03"]
check("лето: LASER / C03 в топ-5, объём < −10%", bool(lz) and lz[0]["объём_к_плану_%"] < -10)
dq = J("data_quality", {})
check("сверка: 14 находок до, 2 после", (dq["до_исправлений"]["находок"], dq["после_исправлений"]["находок"]) == (14, 2))
ev = J("events", {"period": "2025-05", "direction": "DENT"})["события"]
check("events: майские праздники по DENT", len(ev) == 1 and "праздник" in ev[0]["событие"])
check("events: лето, C03 — три простоя лазера", len(J("events", {"period": "2025-07:2025-09", "clinic": "C03"})["события"]) == 3)
check("events: клиника C01 без простоев лазера", J("events", {"period": "2025-07:2025-09", "clinic": "C01"})["события"] == [])
bm = J("pnl", {"period": "2025", "by_month": True})
check("pnl by_month: худший в рублях декабрь, к плану май", (bm["худший_месяц_в_рублях"], bm["худший_месяц_к_плану"]) == ("2025-12", "2025-05"))
check("ошибка периода возвращается текстом", "ошибка" in J("pnl", {"period": "2026"}))
check("неизвестный аргумент не роняет", "ошибка" in J("pnl", {"period": "2025", "foo": 1}))
w = J("write_report", {"title": "t", "markdown": "x", "filename": "../../evil.md"})
check("write_report не выходит из папки", w["сохранено"].startswith("data/agent_reports/") and ".." not in w["сохранено"])
(ROOT / w["сохранено"]).unlink()

# 2. проверка чисел
outs = [tools.call("pnl", {"period": "2025-05"})]
check("верные числа проходят (в т.ч. округление и %)", unverified("Выручка в мае 2025: 47,7 млн ₽ (−9,7%), маржа 2,7 млн ниже плана.", outs) == [])
check("выдуманное число ловится", unverified("Выручка в мае 51,3 млн ₽.", outs) == ["51,3"])
check("рубли с пробелами и номера пунктов", unverified("1. Сумма под риском 120 700 ₽.", [tools.call("data_quality", {})]) == [])


# 3. цикл агента на заглушке модели
class Fake:
    def __init__(self, script):
        self.script, self.i, self.messages = script, 0, self

    def create(self, **kw):
        r = self.script[self.i]
        self.i += 1
        return r


def resp(blocks, stop):
    return NS(content=blocks, stop_reason=stop, usage=NS(input_tokens=100, output_tokens=50))


tu = lambda i, n, a: NS(type="tool_use", id=i, name=n, input=a)  # noqa: E731
tx = lambda t: NS(type="text", text=t)  # noqa: E731
good = Fake([resp([tu("1", "pnl", {"period": "2025-05"})], "tool_use"),
             resp([tu("2", "top_drivers", {"period": "2025-05", "by": "direction"})], "tool_use"),
             resp([tx("В мае 2025 по сети валовая маржа ниже плана на 2,71 млн ₽, главная просадка в стоматологии.")], "end_turn")])
rec = agent.run("Почему упала маржа в мае?", client=good, model="fake", log=False)
check("цикл: 2 вызова, ответ, токены 300/150", len(rec["tool_calls"]) == 2 and rec["usage"] == {"input": 300, "output": 150})
check("цикл: непроверенных чисел нет", rec["unverified_numbers"] == [])
bad = Fake([resp([tu("1", "pnl", {"period": "2025-05"})], "tool_use"), resp([tx("Маржа упала на 8,88 млн ₽.")], "end_turn")])
check("цикл: выдуманное 8,88 помечено", agent.run("q", client=bad, model="fake", log=False)["unverified_numbers"] == ["8,88"])
loop = Fake([resp([tu(str(i), "data_quality", {})], "tool_use") for i in range(agent.MAX_STEPS)])
check("цикл: лимит шагов останавливает", "лимит шагов" in agent.run("q", client=loop, model="fake", log=False)["answer"])
print(f"\n{sum(ok)}/{len(ok)} проверок пройдено")
sys.exit(0 if all(ok) else 1)

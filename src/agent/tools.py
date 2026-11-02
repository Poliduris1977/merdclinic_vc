"""
Инструменты ИИ-агента: только чтение, все числа считает код (pandas), модель их только объясняет.
Источник: data/corrected (после исправлений сверки) и результаты сверки data/output*. Ключ ответов (answer_key) не читается.

Период: "2025" (год), "2025-05" (месяц) или "2025-07:2025-09" (диапазон месяцев). Суммы в млн ₽.
"""
import json
import re
import sys
from functools import lru_cache
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src" / "reports"))
sys.path.insert(0, str(ROOT / "src"))
from build_cashflow import load as load_cash  # noqa: E402
from build_doctors import EVENTS  # noqa: E402
from build_factor import DIRS, effects, prepare, prepare_staff, staff_effects  # noqa: E402
from build_motivation import doctor_table  # noqa: E402

CORR = ROOT / "data" / "corrected"
REPORTS = ROOT / "data" / "agent_reports"
M = 1e6
CLINICS = ("ALL", "C01", "C02", "C03")
BR = [("99", "Прибыль до налога"), ("02", "Амортизация"), ("10", "Запасы материалов"), ("62.02", "Дебиторка ДМС"),
      ("62.01", "Расчёты с пациентами"), ("60", "Долг поставщикам"), ("70", "Долг по зарплате"), ("68", "Долг по НДФЛ"),
      ("69", "Долг по взносам"), ("01", "Капвложения")]


@lru_cache(maxsize=1)
def wh():
    d, _ = prepare()
    e = effects(d)
    se = staff_effects(prepare_staff(e))
    opex = pd.read_csv(CORR / "budget_opex.csv")
    post = pd.read_csv(CORR / "postings_1c.csv", dtype={"debit": str, "credit": str})
    post = post[post.doc_type != "Закрытие месяца"].copy()
    post["period"] = post.date.str[:7]
    dds, osv = load_cash()
    return dict(e=e, se=se, opex=opex[opex.pnl_block != "cost_of_sales"], post=post, dds=dds, osv=osv, dt=doctor_table())


def months(period):
    p = str(period).strip()
    if re.fullmatch(r"2025", p):
        return list(range(1, 13))
    if re.fullmatch(r"2025-(0[1-9]|1[0-2])", p):
        return [int(p[5:])]
    m = re.fullmatch(r"2025-(0[1-9]|1[0-2]):2025-(0[1-9]|1[0-2])", p)
    if m and int(m[1]) <= int(m[2]):
        return list(range(int(m[1]), int(m[2]) + 1))
    raise ValueError('Период: "2025", "2025-05" или "2025-07:2025-09"')


def _clinic(c):
    c = (c or "ALL").upper()
    if c not in CLINICS:
        raise ValueError("Клиника: ALL, C01, C02, C03")
    return c


def _f(df, ms, clinic, direction=None, mcol="period"):
    df = df[df[mcol].astype(str).str[5:7].astype(int).isin(ms)]
    if clinic != "ALL":
        df = df[df.clinic_id == clinic]
    if direction:
        df = df[df.direction_code == direction.upper()]
    return df


def r(x):
    return round(float(x) / M, 2)


def pct(x):
    return round(float(x) * 100, 1)


def _dir(direction):
    if direction and direction.upper() not in DIRS:
        raise ValueError(f"Направление: {', '.join(DIRS)}")
    return direction.upper() if direction else None


# ---------------------------------------------------------------- инструменты
def _pnl_core(ms, clinic):
    w = wh()
    e, se, ox, ps = (_f(w[k], ms, clinic) for k in ("e", "se", "opex", "post"))
    rev_p, rev_f = e.plan_revenue.sum(), e.fact_revenue.sum()
    gm_p = e.plan_margin.sum() - se.plan_costs.sum()
    gm_f = e.fact_margin.sum() - se.fact_costs.sum()
    pr_p = gm_p - ox.plan_amount.sum()
    pr_f = ps[ps.credit == "90.01"].amount.sum() - ps[ps.debit.isin(["20", "26", "44"])].amount.sum()
    return rev_p, rev_f, gm_p, gm_f, pr_p, pr_f


def pnl(period, clinic="ALL", by_month=False):
    ms, clinic = months(period), _clinic(clinic)
    rev_p, rev_f, gm_p, gm_f, pr_p, pr_f = _pnl_core(ms, clinic)
    row = lambda p, f: {"план": r(p), "факт": r(f), "отклонение": r(f - p)}  # noqa: E731
    out = {"период": str(period), "клиника": clinic, "единицы": "млн ₽",
           "выручка": row(rev_p, rev_f), "выручка_отклонение_%": pct(rev_f / rev_p - 1),
           "валовая_маржа": row(gm_p, gm_f),
           "валовая_маржа_%_выручки": {"план": pct(gm_p / rev_p), "факт": pct(gm_f / rev_f)},
           "прибыль_до_налога": row(pr_p, pr_f)}
    if by_month and len(ms) > 1:
        rows = []
        for m in ms:
            a, b, c, d, _, _ = _pnl_core([m], clinic)
            rows.append({"месяц": f"2025-{m:02d}", "выручка_откл": r(b - a), "выручка_откл_%": pct(b / a - 1),
                         "валовая_откл": r(d - c), "валовая_откл_%_плана": pct((d - c) / c)})
        out["по_месяцам"] = rows
        out["худший_месяц_в_рублях"] = min(rows, key=lambda x: x["валовая_откл"])["месяц"]
        out["худший_месяц_к_плану"] = min(rows, key=lambda x: x["валовая_откл_%_плана"])["месяц"]
    return out


FACTORS = [("volume", "Объём"), ("mix", "Структура услуг"), ("price_list", "Цена: прайс"), ("discount", "Цена: скидки"),
           ("materials", "Материалы"), ("labs", "Лаборатории"), ("doc_vol", "Врачи: объём"), ("doc_price", "Врачи: цена и скидки"),
           ("doc_grade", "Врачи: грейды и доля"), ("nurse_eff", "Медперсонал"), ("depr_eff", "Амортизация")]


def margin_bridge(period, clinic="ALL", direction=None):
    ms, clinic, direction, w = months(period), _clinic(clinic), _dir(direction), wh()
    e, se = _f(w["e"], ms, clinic, direction), _f(w["se"], ms, clinic, direction)
    gm_p = e.plan_margin.sum() - se.plan_costs.sum()
    gm_f = e.fact_margin.sum() - se.fact_costs.sum()
    vals = {**e[["volume", "mix", "price_list", "discount", "materials", "labs"]].sum().to_dict(),
            **se[["doc_vol", "doc_price", "doc_grade", "nurse_eff", "depr_eff"]].sum().to_dict()}
    out = {"период": str(period), "клиника": clinic, "направление": direction or "все", "единицы": "млн ₽",
           "валовая_маржа": {"план": r(gm_p), "факт": r(gm_f), "отклонение": r(gm_f - gm_p)},
           "факторы": [{"фактор": lbl, "эффект": r(vals[k])} for k, lbl in FACTORS]}
    if direction:
        out["примечание"] = "Медперсонал и амортизация заданы на уровне клиники и по направлениям не распределяются (в срезе 0)"
    return out


def top_drivers(period, clinic="ALL", direction=None, by="direction", n=5):
    ms, clinic, direction, w = months(period), _clinic(clinic), _dir(direction), wh()
    e = _f(w["e"], ms, clinic, direction)
    keys = {"direction": ["direction_code"], "clinic": ["clinic_id"], "service": ["service_name"],
            "direction_clinic": ["direction_code", "clinic_id"]}
    if by not in keys:
        raise ValueError("by: direction, clinic, service или direction_clinic")
    e = e.assign(price=e.price_list + e.discount, dev=e.fact_margin - e.plan_margin)
    g = e.groupby(keys[by]).agg(dev=("dev", "sum"), volume=("volume", "sum"), mix=("mix", "sum"), price=("price", "sum"),
                                materials=("materials", "sum"), labs=("labs", "sum"), qf=("q_f_pp", "sum"), qp=("q_p_pp", "sum"),
                                rev_p=("plan_revenue", "sum"), rev_f=("fact_revenue", "sum")).reset_index()
    g["volume_pct"] = pct_s(g.qf / g.qp - 1)
    g["rev_dev_pct"] = pct_s(g.rev_f / g.rev_p - 1)
    g = g.sort_values("dev").head(int(n))
    rows = []
    for x in g.itertuples():
        key = " / ".join(str(getattr(x, k)) for k in ["direction_code", "clinic_id", "service_name"] if hasattr(x, k))
        rows.append({"срез": key, "отклонение_прямой_маржи": r(x.dev), "объём": r(x.volume), "структура": r(x.mix),
                     "цена": r(x.price), "материалы": r(x.materials), "лаборатории": r(x.labs),
                     "объём_к_плану_%": x.volume_pct, "выручка_к_плану_%": x.rev_dev_pct})
    return {"период": str(period), "клиника": clinic, "направление": direction or "все", "разрез": by, "единицы": "млн ₽",
            "прямая_маржа_до_оплаты_врачей": "отклонения в этом инструменте считаются до оплаты труда (прямая маржа)",
            "худшие": rows, "направления": DIRS}


def pct_s(s):
    return (s * 100).round(1)


def cashflow(period, clinic="ALL"):
    ms, clinic, w = months(period), _clinic(clinic), wh()
    dds = w["dds"][w["dds"].m.isin(ms)]
    osv = w["osv"][w["osv"].m.isin(ms)]
    if clinic != "ALL":
        dds, osv = dds[dds.clinic_id == clinic], osv[osv.clinic_id == clinic]
    op = dds[dds.activity == "operating"].signed_amount.sum()
    inv = dds[dds.activity == "investing"].signed_amount.sum()
    a51 = osv[osv.acct == "сч. 51"]
    first, last = min(ms), max(ms)
    j = osv[osv.m == first].groupby("acct").open_net.sum()
    c = osv[osv.m == last].groupby("acct").close_net.sum()
    bridge = [{"статья": lbl, "эффект": r(-(c.get(f"сч. {a}", 0) - j.get(f"сч. {a}", 0)))} for a, lbl in BR]
    return {"период": str(period), "клиника": clinic, "единицы": "млн ₽",
            "операционный_поток": r(op), "инвестиционный_поток": r(inv), "чистое_изменение": r(op + inv),
            "остаток_на_начало": r(a51[a51.m == first].open_net.sum()), "остаток_на_конец": r(a51[a51.m == last].close_net.sum()),
            "мост_прибыль_в_деньги": bridge}


def doctors(clinic="ALL", direction=None):
    clinic, direction, dt = _clinic(clinic), _dir(direction), wh()["dt"]
    if clinic != "ALL":
        dt = dt[dt.clinic_id == clinic]
    if direction:
        dt = dt[dt.direction_code == direction]
    dt = dt.assign(margin=dt.revenue - dt.norm_cost - dt.pay_total, ret_n=dt.cohort_n * dt.returned)
    g = dt.groupby("grade").agg(n=("doctor_id", "count"), rev=("revenue", "sum"), pay=("pay_total", "sum"), mg=("margin", "sum"))
    t = dt.groupby("direction_code").agg(c=("cohort_n", "sum"), rr=("ret_n", "sum"))
    return {"клиника": clinic, "направление": direction or "все", "единицы": "млн ₽", "врачей": int(len(dt)),
            "грейды": [{"грейд": k, "врачей": int(x.n), "выручка": r(x.rev), "оплата_с_взносами": r(x.pay),
                        "доля_оплаты_%": pct(x.pay / x.rev), "маржа_после_оплаты_%": pct(x.mg / x.rev)} for k, x in g.iterrows()],
            "повторные_визиты_90д_%": {DIRS[k]: pct(x.rr / x.c) for k, x in t.iterrows() if x.c},
            "примечание": "Малая выборка: 6-12 врачей на направление; повторные визиты нельзя сравнивать между направлениями"}


def data_quality():
    out = {}
    for name, p in (("до_исправлений", "output"), ("после_исправлений", "output_corrected")):
        f = pd.read_csv(ROOT / "data" / p / "findings.csv", encoding="utf-8-sig")
        f = f[f.is_root_cause.astype(str).str.lower() == "true"]
        out[name] = {"находок": int(len(f)), "сумма_под_риском_руб": int(f.amount_rub.abs().sum()),
                     "по_типам": {k: int(v) for k, v in f.error_type.value_counts().items()}}
    out["примечание"] = ("После исправлений остаются 2 визита без врача: исправление намеренно не применено, "
                         "в отчётах это отдельный остаток. Три расхождения выручки исправлены эвристикой, подтвердить бухгалтером.")
    return out


def events(period="2025", clinic="ALL", direction=None):
    """Известные разовые события (вход, как на листе «Сезонность» книги DoctorAnalytics). Причину события данные не доказывают."""
    ms, clinic, direction = months(period), _clinic(clinic), _dir(direction)
    out = []
    for per, d, c, what in EVENTS:
        if int(per[5:]) not in ms or (direction and d != direction) or (clinic != "ALL" and c not in ("Все", clinic)):
            continue
        out.append({"период": per, "направление": d, "клиники": "все" if c == "Все" else c, "событие": what})
    return {"период": str(period), "клиника": clinic, "события": out,
            "примечание": "Календарь известных событий; влияние на числа оценивай по top_drivers, а не по самому факту события"}


def write_report(title, markdown, filename):
    REPORTS.mkdir(parents=True, exist_ok=True)
    name = re.sub(r"[^\w\-]+", "_", Path(filename).stem, flags=re.U)[:60] or "report"
    path = REPORTS / f"{name}.md"
    path.write_text(f"# {title}\n\n{markdown}\n", encoding="utf-8")
    return {"сохранено": path.relative_to(ROOT).as_posix()}


# ---------------------------------------------------------------- схемы для модели
_P = {"period": {"type": "string", "description": 'Период: "2025", "2025-05" или "2025-07:2025-09"'},
      "clinic": {"type": "string", "enum": list(CLINICS), "description": "Клиника, ALL = вся сеть"},
      "direction": {"type": "string", "enum": list(DIRS), "description": "Код направления (необязательно)"}}


def _t(name, desc, props, req):
    return {"name": name, "description": desc, "input_schema": {"type": "object", "properties": props, "required": req}}


TOOLS = [
    _t("pnl", "План и факт: выручка, валовая маржа (в млн ₽ и % выручки), прибыль до налога за период и клинику. "
              "by_month=true добавляет разбивку по месяцам и худший месяц (в рублях и к плану): обязательно для вопросов про «худший», «лучший», динамику.",
       {"period": _P["period"], "clinic": _P["clinic"], "by_month": {"type": "boolean"}}, ["period"]),
    _t("margin_bridge", "Факторный мост валовой маржи план → факт: 11 факторов (объём, структура, цена, материалы, врачи, персонал, амортизация).",
       {"period": _P["period"], "clinic": _P["clinic"], "direction": _P["direction"]}, ["period"]),
    _t("top_drivers", "Где просела прямая маржа: худшие срезы с разложением на объём, структуру, цену, материалы. Используй для вопросов «почему упала маржа».",
       {"period": _P["period"], "clinic": _P["clinic"], "direction": _P["direction"],
        "by": {"type": "string", "enum": ["direction", "clinic", "service", "direction_clinic"], "description": "Разрез ранжирования"}}, ["period"]),
    _t("cashflow", "Денежный поток за период: операционный, инвестиционный, остатки и мост от прибыли до денег.",
       {"period": _P["period"], "clinic": _P["clinic"]}, ["period"]),
    _t("doctors", "Врачи: грейды (выручка, доля оплаты, маржа), повторные визиты за 90 дней по направлениям.",
       {"clinic": _P["clinic"], "direction": _P["direction"]}, []),
    _t("events", "Календарь известных разовых событий (праздники, простои оборудования) по периоду, клинике и направлению. "
                 "Причину просадки называй фактом только если событие есть здесь; иначе это гипотеза.",
       {"period": _P["period"], "clinic": _P["clinic"], "direction": _P["direction"]}, ["period"]),
    _t("data_quality", "Результат сверки 1С и ERP: количество находок и сумма под риском до и после исправлений.", {}, []),
    _t("write_report", "Сохранить готовый отчёт или резюме в markdown-файл. Вызывать только когда человек просит отчёт или резюме.",
       {"title": {"type": "string"}, "markdown": {"type": "string", "description": "Текст отчёта, только числа из инструментов"},
        "filename": {"type": "string", "description": "Имя файла без пути"}}, ["title", "markdown", "filename"]),
]
REGISTRY = {"pnl": pnl, "margin_bridge": margin_bridge, "top_drivers": top_drivers, "cashflow": cashflow,
            "doctors": doctors, "events": events, "data_quality": data_quality, "write_report": write_report}


def call(name, args):
    """Вызов инструмента: ошибки возвращаются модели текстом, а не роняют агента."""
    try:
        return json.dumps(REGISTRY[name](**args), ensure_ascii=False)
    except Exception as ex:  # noqa: BLE001
        return json.dumps({"ошибка": str(ex)}, ensure_ascii=False)

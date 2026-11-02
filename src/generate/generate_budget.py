"""
Шаг 4 генератора: бюджет (план) на 2025 год.

План ЗАВЫШЕН: факт ниже плана примерно на 3-5% по выручке. Бюджет строится
от факта «в обратную сторону» с заложенными причинами отклонений, которые
потом должны найти план-факт и факторный анализ маржи.

Выход (data/raw/):
  budget_revenue.csv - план по месяцам / клиникам / услугам:
                       количество, плановая цена, плановая себестоимость единицы
  budget_opex.csv    - план расходов по клиникам и статьям
Выход (data/answer_key/):
  planted_findings.csv - ключ ответов: что именно заложено (НЕ подавать агенту)

Заложенные причины отклонений:
  1. Стоматология, май: план не учёл майские праздники -> провал объёма (все клиники)
  2. Лазерная коррекция, клиника C03, июль-сентябрь: простой оборудования -> объём
  3. Косметология: план закладывал индексацию прайса +8% с апреля, факт +4% -> цена
  4. Везде: план по цене без скидок, факт со скидками -> небольшой ценовой эффект
  5. Себестоимость: план закладывал рост цен на материалы 0.3%/мес, факт растёт быстрее -> затраты
  + фоновый недобор объёма 1.5-3.5% в каждой ячейке (оптимизм плана)

Запуск:  python src/generate/generate_budget.py
"""
import random
from datetime import date
from pathlib import Path

import pandas as pd

from generate_1c import ACQUIRING, ADMIN, CONTRIB, DEPR, MARKETING_PCT, NURSES, UTIL_PCT_OF_RENT

SEED = 11
YEAR = 2025
RAW = Path(__file__).resolve().parents[2] / "data" / "raw"
KEY = RAW.parent / "answer_key"
rng = random.Random(SEED)

BASE_GAP = (0.015, 0.035)           # фоновый недобор объёма относительно плана
VOLUME_EVENTS = [                    # (направление, клиника или None, месяцы, доп. недобор)
    ("DENT", None, [5], 0.12),
    ("LASER", "C03", [7, 8, 9], 0.12),
]
COSM_PLAN_INDEX = 1.08               # план: +8% с апреля (факт +4%)
PLAN_MAT_INFLATION = 0.003           # план: рост цен на материалы в месяц
PLAN_EXT_INFLATION = 0.002
SELF_SHARE = 0.88                    # доля оплат через эквайринг, заложенная в план


def round100(x):
    return int(round(x / 100.0) * 100)


def main():
    clinics = pd.read_csv(RAW / "clinics.csv")
    services = pd.read_csv(RAW / "services.csv")
    doctors = pd.read_csv(RAW / "doctors.csv")
    visits = pd.read_csv(RAW / "visits.csv", parse_dates=["visit_date"])
    items = pd.read_csv(RAW / "visit_items.csv")
    history = pd.read_csv(RAW / "price_history.csv")

    # --- факт по месяцам / клиникам / услугам --------------------------------
    im = items.merge(visits[["visit_id", "visit_date", "clinic_id"]], on="visit_id")
    im["month"] = im.visit_date.dt.month
    fact = im.groupby(["month", "clinic_id", "service_id"]).agg(
        fact_qty=("qty", "sum"), fact_revenue=("revenue", "sum")).reset_index()
    fact = fact.merge(services[["service_id", "direction_code", "list_price",
                                "material_cost", "external_cost"]], on="service_id")

    # --- плановая цена: прайс по плану на 1-е число месяца --------------------
    hist = history.copy()
    hist["valid_from"] = pd.to_datetime(hist.valid_from)

    def plan_price(sid, direction, m):
        h = hist[(hist.service_id == sid) & (hist.valid_from <= pd.Timestamp(YEAR, m, 1))]
        actual = int(h.sort_values("valid_from").iloc[-1].list_price)
        if direction == "COSM":
            base = int(services.loc[services.service_id == sid, "list_price"].iloc[0])
            return round100(base * COSM_PLAN_INDEX) if m >= 4 else base
        return actual

    # --- недобор объёма по ячейкам -------------------------------------------
    gap = {}
    for (m, c, d), _ in fact.groupby(["month", "clinic_id", "direction_code"]):
        g = rng.uniform(*BASE_GAP)
        for ev_dir, ev_clinic, ev_months, extra in VOLUME_EVENTS:
            if d == ev_dir and m in ev_months and (ev_clinic is None or ev_clinic == c):
                g += extra
        gap[(m, c, d)] = g

    rows = []
    for r in fact.itertuples():
        g = gap[(r.month, r.clinic_id, r.direction_code)]
        plan_qty = round(r.fact_qty / (1 - g))
        pp = plan_price(r.service_id, r.direction_code, r.month)
        rows.append(dict(
            period=f"{YEAR}-{r.month:02d}", clinic_id=r.clinic_id,
            direction_code=r.direction_code, service_id=r.service_id,
            plan_qty=plan_qty, plan_price=pp, plan_revenue=plan_qty * pp,
            plan_unit_material=round(r.material_cost * (1 + PLAN_MAT_INFLATION * (r.month - 1)), 2),
            plan_unit_external=round(r.external_cost * (1 + PLAN_EXT_INFLATION * (r.month - 1)), 2),
        ))
    bud = pd.DataFrame(rows).sort_values(["period", "clinic_id", "service_id"])

    # --- план расходов --------------------------------------------------------
    pct = doctors.groupby(["clinic_id", "direction_code"]).pay_pct.mean().to_dict()
    rev_cd = bud.groupby(["period", "clinic_id", "direction_code"]).plan_revenue.sum()
    opex = []
    for (period, cid), g in rev_cd.groupby(level=[0, 1]):
        rev = g.sum()
        doc = sum(v * pct.get((cid, d), 0.33) for (_, _, d), v in g.items())
        rent = float(clinics.loc[clinics.clinic_id == cid, "rent_month"].iloc[0])
        lines = [
            ("Оплата труда врачей", "cost_of_sales", doc),
            ("Оплата труда медперсонала", "cost_of_sales", NURSES[cid]),
            ("Страховые взносы (себестоимость)", "cost_of_sales", (doc + NURSES[cid]) * CONTRIB),
            ("Амортизация", "cost_of_sales", DEPR[cid]),
            ("Оплата труда административного персонала", "management", ADMIN[cid]),
            ("Страховые взносы (администрация)", "management", ADMIN[cid] * CONTRIB),
            ("Аренда", "management", rent),
            ("Коммунальные услуги", "management", rent * UTIL_PCT_OF_RENT),
            ("Комиссия эквайринга", "management", rev * SELF_SHARE * ACQUIRING),
            ("Маркетинг", "commercial", rev * MARKETING_PCT),
        ]
        for article, block, amt in lines:
            opex.append(dict(period=period, clinic_id=cid, article=article,
                             pnl_block=block, plan_amount=round(amt, 2)))
    opex = pd.DataFrame(opex)

    # --- ключ ответов ----------------------------------------------------------
    key = pd.DataFrame([
        dict(finding_id="B1", period="2025-05", clinic_id="all", direction_code="DENT",
             type="volume", description="План не учёл майские праздники: недобор объёма в стоматологии сверх фонового (~12 п.п.)"),
        dict(finding_id="B2", period="2025-07..2025-09", clinic_id="C03", direction_code="LASER",
             type="volume", description="Простой оборудования: недобор объёма лазерной коррекции в C03 (~12 п.п.)"),
        dict(finding_id="B3", period="2025-04..2025-12", clinic_id="all", direction_code="COSM",
             type="price", description="План закладывал индексацию +8% с апреля, факт +4%: ценовое отклонение косметологии"),
        dict(finding_id="B4", period="2025-01..2025-12", clinic_id="all", direction_code="all",
             type="price", description="План по прайсу без скидок: фактическая цена ниже из-за скидок"),
        dict(finding_id="B5", period="2025-01..2025-12", clinic_id="all", direction_code="all",
             type="cost", description="Плановый рост цен на материалы 0.3%/мес, факт 0.5%/мес плюс шум: перерасход материалов"),
        dict(finding_id="B6", period="2025-01..2025-12", clinic_id="all", direction_code="all",
             type="volume", description="Фоновый оптимизм плана: недобор объёма 1.5-3.5% в каждой ячейке"),
    ])

    KEY.mkdir(parents=True, exist_ok=True)
    bud.to_csv(RAW / "budget_revenue.csv", index=False, encoding="utf-8-sig")
    opex.to_csv(RAW / "budget_opex.csv", index=False, encoding="utf-8-sig")
    key.to_csv(KEY / "planted_findings.csv", index=False, encoding="utf-8-sig")

    # --- сводка -------------------------------------------------------------
    plan_rev = bud.plan_revenue.sum()
    fact_rev = fact.fact_revenue.sum()
    print(f"Выручка: план {plan_rev / 1e6:,.1f} млн, факт {fact_rev / 1e6:,.1f} млн, "
          f"отклонение {(fact_rev / plan_rev - 1):.1%}")

    pb = bud.groupby("direction_code").plan_revenue.sum()
    fb = fact.groupby("direction_code").fact_revenue.sum()
    print("\nОтклонение по направлениям:")
    print(pd.DataFrame({"план, млн": pb / 1e6, "факт, млн": fb / 1e6,
                        "откл.": fb / pb - 1}).round(3).to_string())

    pm = bud.assign(m=bud.period.str[-2:]).groupby("m").plan_revenue.sum()
    fm = fact.groupby("month").fact_revenue.sum()
    fm.index = [f"{i:02d}" for i in fm.index]
    print("\nОтклонение по месяцам:")
    print((fm / pm - 1).round(3).to_string())

    dent_may = (fact[(fact.month == 5) & (fact.direction_code == "DENT")].fact_revenue.sum() /
                bud[(bud.period == "2025-05") & (bud.direction_code == "DENT")].plan_revenue.sum() - 1)
    laser = (fact[(fact.month.isin([7, 8, 9])) & (fact.clinic_id == "C03") &
                  (fact.direction_code == "LASER")].fact_revenue.sum() /
             bud[(bud.period.isin(["2025-07", "2025-08", "2025-09"])) & (bud.clinic_id == "C03") &
                 (bud.direction_code == "LASER")].plan_revenue.sum() - 1)
    print(f"\nЗаложенные события: стоматология май {dent_may:.1%}, лазер C03 июль-сентябрь {laser:.1%}")


if __name__ == "__main__":
    main()

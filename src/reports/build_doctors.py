"""
Шаг 12: аналитика врачей, 2025 (Excel).

Листы:
  Врачи       выручка, визиты, чек, маржа после ФОТ, повторные визиты по каждому врачу (формулы по «Данным»)
  Грейды      чек, доля оплаты и маржа по направлениям и грейдам (формулы): мост к факторному анализу
  Пациенты    повторные визиты за 90 дней и LTV-90 по источникам и направлениям (значения, считает скрипт)
  Сезонность  индекс месяца к среднему по направлению с отметкой известных разовых событий (формулы)
  Данные      «месяц x врач» (входные числа и формулы маржи)

Решения (демо):
  * врачи сравниваются внутри направления и грейда: между направлениями чек различается в разы;
  * материалы и лаборатории - по нормативу услуг (количество x норма единицы), без разнесения факта 1С;
  * оплата врача - факт из payroll_1c (начислено + взносы); ставка взносов в демо - допущение (см. факторный анализ);
  * повторный визит: пациент вернулся в то же направление в течение 90 дней (учитываются только первые визиты
    к врачу до 02.10.2025, чтобы 90-дневное окно целиком лежало в 2025 году);
  * LTV-90: выручка нового пациента 2025 года за 90 дней от первого визита (когорты с первым визитом до 02.10.2025);
    данные есть только за 2025 год, поэтому длиннее 90 дней LTV не считается;
  * прогноз не строится: по ряду из одного года сезонность не отделить от разовых событий.

Вход:  data/corrected/
Выход: data/reports/DoctorAnalytics_2025.xlsx

Запуск:  python src/reports/build_doctors.py
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from build_factor import (BLACK, BLUE, BOLD, DIRS, F, GREEN, GREEN_LINK, INPUT_FILL, NOTE, NUM, PCT,  # noqa: E402
                          RED, SEC_FILL, TITLE, TOP, head, sign_format)
from common.recalc import recalculate  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "reports"
YEAR = 2025
CLINICS = ["C01", "C02", "C03"]
GRADES = ["lead", "senior", "middle", "junior"]
NODOC = "—"
WINDOW = 90
COHORT_END = pd.Timestamp("2025-10-02")     # + 90 дней = 31.12.2025
RUB = "#,##0;(#,##0);-"
EVENTS = [  # известные разовые события (вход, редактируется на листе «Сезонность»)
    ("2025-05", "DENT", "Все", "Майские праздники: сокращённый график работы"),
    ("2025-07", "LASER", "C03", "Простой лазерного оборудования"),
    ("2025-08", "LASER", "C03", "Простой лазерного оборудования"),
    ("2025-09", "LASER", "C03", "Простой лазерного оборудования"),
]


# ------------------------------------------------------------------ данные
def prepare():
    visits = pd.read_csv(CORR / "visits.csv")
    items = pd.read_csv(CORR / "visit_items.csv")
    services = pd.read_csv(CORR / "services.csv")
    doctors = pd.read_csv(CORR / "doctors.csv")
    pay = pd.read_csv(CORR / "payroll_1c.csv")
    patients = pd.read_csv(CORR / "patients.csv")
    budget = pd.read_csv(CORR / "budget_revenue.csv")
    plan = (budget.assign(plan_rev=budget.plan_qty * budget.plan_price)
            .groupby(["period", "direction_code"], as_index=False).plan_rev.sum())

    x = items.merge(services[["service_id", "material_cost", "external_cost"]], on="service_id")
    x["cost"] = x.qty * (x.material_cost + x.external_cost)
    vr = x.groupby("visit_id", as_index=False).agg(rev=("revenue", "sum"), cost=("cost", "sum"))
    done = visits[visits.status == "completed"].merge(vr, on="visit_id", how="left").fillna({"rev": 0, "cost": 0})
    done["period"] = done.visit_date.str[:7]
    done["doctor_id"] = done.doctor_id.fillna(NODOC)

    # месяц x врач (и строки «Без врача» по клиникам и направлениям)
    g = done.groupby(["period", "doctor_id", "clinic_id", "direction_code"], as_index=False).agg(
        visits=("visit_id", "count"), revenue=("rev", "sum"), norm_cost=("cost", "sum"))
    dm = doctors[["doctor_id", "doctor_name", "grade"]]
    g = g.merge(dm, on="doctor_id", how="left")
    g["doctor_name"] = g.doctor_name.fillna("Без врача")
    g["grade"] = g.grade.fillna(NODOC)
    p = pay[pay.role == "doctor"].assign(pay_total=lambda d: d.gross + d.contributions)
    p = p.groupby(["period", "employee_id"], as_index=False).pay_total.sum().rename(columns={"employee_id": "doctor_id"})
    g = g.merge(p, on=["period", "doctor_id"], how="left").fillna({"pay_total": 0.0})
    miss = p.merge(g[["period", "doctor_id"]].drop_duplicates(), on=["period", "doctor_id"], how="left", indicator=True)
    if (miss._merge == "left_only").any() and miss[miss._merge == "left_only"].pay_total.abs().sum() > 0:
        raise ValueError("ФОТ врача за месяц без визитов: нужен разбор")
    g = g.sort_values(["period", "direction_code", "clinic_id", "doctor_id"]).reset_index(drop=True)
    data = g[["period", "doctor_id", "doctor_name", "direction_code", "clinic_id", "grade",
              "visits", "revenue", "norm_cost", "pay_total"]].copy()
    for c in ("revenue", "norm_cost", "pay_total"):
        data[c] = data[c].round(2)

    # повторные визиты за 90 дней: вернулся в то же направление
    dd = done[["patient_id", "direction_code", "visit_date"]].drop_duplicates()
    dd["d"] = pd.to_datetime(dd.visit_date)
    dd = dd.sort_values(["patient_id", "direction_code", "d"])
    dd["next_d"] = dd.groupby(["patient_id", "direction_code"]).d.shift(-1)
    dd["returned"] = (dd.next_d - dd.d).dt.days.le(WINDOW)
    done["d"] = pd.to_datetime(done.visit_date)
    done = done.merge(dd[["patient_id", "direction_code", "d", "returned"]], on=["patient_id", "direction_code", "d"], how="left")
    first_doc = (done[(done.doctor_id != NODOC) & (done.d <= COHORT_END)]
                 .sort_values(["d", "visit_id"]).drop_duplicates(["patient_id", "doctor_id"]))
    ret = first_doc.groupby("doctor_id").agg(cohort_n=("patient_id", "count"), returned=("returned", "mean")).reset_index()

    # LTV-90 новых пациентов 2025
    new = patients[patients.first_visit_date >= f"{YEAR}-01-01"][["patient_id", "source"]]
    fv = (done.sort_values(["d", "visit_id"]).drop_duplicates("patient_id")[["patient_id", "d", "direction_code"]]
          .rename(columns={"d": "first_d", "direction_code": "first_dir"}))
    co = new.merge(fv, on="patient_id")
    co = co[co.first_d <= COHORT_END]
    vv = done[["patient_id", "d", "rev"]].merge(co[["patient_id", "first_d"]], on="patient_id")
    vv = vv[(vv.d >= vv.first_d) & ((vv.d - vv.first_d).dt.days <= WINDOW)]
    ltv = vv.groupby("patient_id").rev.sum().rename("ltv")
    nxt = vv.assign(later=vv.d > vv.first_d).groupby("patient_id").later.max().rename("came_back")
    co = co.join(ltv, on="patient_id").join(nxt, on="patient_id").fillna({"ltv": 0.0, "came_back": False})
    return data, doctors, ret, co, plan


# ------------------------------------------------------------------ книга
def build(data, doctors, ret, co, plan):
    wb = Workbook()
    nd = len(data) + 1
    D = lambda c: f"'Данные'!${c}$2:${c}${nd}"  # noqa: E731

    # ---------------------------------------------------------------- Данные
    wd = wb.active
    wd.title = "Данные"
    hdr = ["Период", "Врач", "ФИО", "Направление", "Клиника", "Грейд", "Визитов", "Выручка",
           "Материалы и лаборатории (норматив)", "Оплата врача с взносами, факт", "Маржа после ФОТ"]
    for i, h in enumerate(hdr, 1):
        wd.cell(1, i, h)
    head(wd, 1, 1, len(hdr))
    wd.row_dimensions[1].height = 42
    for r, row in enumerate(data.itertuples(index=False), 2):
        for c, v in enumerate(row, 1):
            cell = wd.cell(r, c, v)
            cell.font = BLUE if c >= 7 else BLACK
            if c >= 8:
                cell.number_format = "#,##0.00"
        wd.cell(r, 11, f"=H{r}-I{r}-J{r}").font = BLACK
        wd.cell(r, 11).number_format = "#,##0.00"
    for i, w in enumerate([9, 8, 16, 12, 9, 8, 9, 14, 16, 16, 14], 1):
        wd.column_dimensions[col(i)].width = w
    wd.freeze_panes = "D2"
    wd.auto_filter.ref = f"A1:K{nd}"

    # ---------------------------------------------------------------- Врачи
    ws = wb.create_sheet("Врачи", 0)
    ws["A1"], ws["A1"].font = "Врачи: выручка, чек и маржа после оплаты труда, 2025", TITLE
    ws["A2"] = ("Сравнивайте врачей только внутри направления и специальности: чек различается между направлениями в разы, а внутри "
                "направления зависит от специальности (например, ортопед и терапевт). "
                "Маржа после ФОТ = выручка − материалы и лаборатории (норматив) − оплата врача с взносами.")
    ws["A3"] = ("Малая выборка: у направления 1-8 врачей, выводы о конкретном враче по одному году делать нельзя. "
                "Взносы: допущение демо (плоские 22%).")
    ws["A2"].font = ws["A3"].font = NOTE
    h = ["Врач", "ФИО", "Направление: специальность", "Клиника", "Грейд", "Визитов", "Выручка, тыс. ₽", "Чек, ₽",
         "Чек к среднему по направлению", "Оплата врача с взносами, тыс. ₽", "Доля оплаты с взносами", "Материалы и лаб., тыс. ₽",
         "Маржа после ФОТ, тыс. ₽", "Маржа, % выручки", "Пациентов в когорте", "Вернулись за 90 дней"]
    for i, t in enumerate(h, 1):
        ws.cell(5, i, t)
    head(ws, 5, 1, len(h))
    ws.row_dimensions[5].height = 44
    order = {d: i for i, d in enumerate(DIRS)}
    gorder = {g: i for i, g in enumerate(GRADES)}
    doc = doctors.assign(o=doctors.direction_code.map(order), g=doctors.grade.map(gorder)).sort_values(
        ["o", "g", "clinic_id", "doctor_id"])
    retd = ret.set_index("doctor_id")
    r = 6
    first_r = r
    for x in doc.itertuples():
        ws.cell(r, 1, x.doctor_id).font = BLACK
        ws.cell(r, 2, x.doctor_name).font = BLACK
        ws.cell(r, 3, f"{DIRS[x.direction_code]}: {x.specialty}").font = BLACK
        ws.cell(r, 4, x.clinic_id).font = BLACK
        ws.cell(r, 5, x.grade).font = BLACK
        ws.cell(r, 18, x.direction_code).font = Font(name=F, size=8, color="808080")
        ws.cell(r, 15, int(retd.cohort_n.get(x.doctor_id, 0))).font = BLACK
        ws.cell(r, 16, float(retd.returned.get(x.doctor_id, 0)) if x.doctor_id in retd.index else 0).font = BLACK
        r += 1
    last_doc = r - 1
    ws.cell(r, 1, NODOC).font = BLACK
    ws.cell(r, 2, "Без врача (остаток сверки)").font = BLACK
    nodoc_r = r
    r += 1
    tot = r
    for rr in range(first_r, nodoc_r + 1):
        key = f"$A{rr}"
        ws.cell(rr, 6, f"=SUMIFS({D('G')},{D('B')},{key})")
        ws.cell(rr, 7, f"=SUMIFS({D('H')},{D('B')},{key})")
        ws.cell(rr, 8, f"=IF(F{rr}=0,0,G{rr}/F{rr})")
        if rr != nodoc_r:
            ws.cell(rr, 9, f"=IF(H{rr}=0,0,H{rr}/(SUMIFS({D('H')},{D('D')},$R{rr})/SUMIFS({D('G')},{D('D')},$R{rr})))")
        ws.cell(rr, 10, f"=SUMIFS({D('J')},{D('B')},{key})")
        ws.cell(rr, 11, f"=IF(G{rr}=0,0,J{rr}/G{rr})")
        ws.cell(rr, 12, f"=SUMIFS({D('I')},{D('B')},{key})")
        ws.cell(rr, 13, f"=G{rr}-L{rr}-J{rr}")
        ws.cell(rr, 14, f"=IF(G{rr}=0,0,M{rr}/G{rr})")
        for c in range(6, 15):
            ws.cell(rr, c).font = GREEN_LINK if c in (6, 7, 10, 12) else BLACK
        ws.cell(rr, 6).number_format = "#,##0"
        ws.cell(rr, 8).number_format = RUB
        ws.cell(rr, 9).number_format = ws.cell(rr, 11).number_format = ws.cell(rr, 14).number_format = PCT
        for c in (7, 10, 12, 13):
            ws.cell(rr, c).number_format = NUM
        ws.cell(rr, 16).number_format = PCT
        ws.cell(rr, 15).number_format = "#,##0"
    ws.cell(tot, 1, "Итого").font = BOLD
    for c in (6, 7, 10, 12, 13):
        L = col(c)
        ws.cell(tot, c, f"=SUM({L}{first_r}:{L}{nodoc_r})").font = BOLD
        ws.cell(tot, c).number_format = NUM if c != 6 else "#,##0"
    ws.cell(tot, 8, f"=IF(F{tot}=0,0,G{tot}/F{tot})").number_format = RUB
    ws.cell(tot, 11, f"=IF(G{tot}=0,0,J{tot}/G{tot})").number_format = PCT
    ws.cell(tot, 14, f"=IF(G{tot}=0,0,M{tot}/G{tot})").number_format = PCT
    for c in (8, 11, 14):
        ws.cell(tot, c).font = BOLD
    for c in range(1, 17):
        ws.cell(tot, c).border = TOP
    ws.cell(tot + 1, 1, "Контроль: итог листа минус итог «Данных», выручка").font = NOTE
    ws.cell(tot + 1, 7, f"=G{tot}-SUM({D('H')})").font = NOTE
    ws.cell(tot + 1, 7).number_format = "#,##0.00"
    ws.cell(tot + 2, 1, "Пациентов в когорте и «Вернулись за 90 дней»: значения, рассчитаны скриптом (первый визит к врачу до 02.10.2025; "
                        "вернулся = новый визит пациента в то же направление в течение 90 дней).").font = NOTE
    sign_format(ws, f"M{first_r}:M{nodoc_r}")
    ws.conditional_formatting.add(f"I{first_r}:I{last_doc}", ColorScaleRule(
        start_type="num", start_value=0.8, start_color="F8CBAD", mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="num", end_value=1.2, end_color="C6E0B4"))
    for i, w in enumerate([8, 18, 40, 8, 8, 10, 13, 10, 14, 13, 12, 14, 14, 11, 12, 12], 1):
        ws.column_dimensions[col(i)].width = w
    ws.column_dimensions["R"].width = 7
    ws.freeze_panes = "C6"
    ws.auto_filter.ref = f"A5:P{nodoc_r}"
    doc_sheet_rng = lambda c: f"'Врачи'!${c}${first_r}:${c}${last_doc}"  # noqa: E731

    # ---------------------------------------------------------------- Грейды
    wg = wb.create_sheet("Грейды", 1)
    wg["A1"], wg["A1"].font = "Грейды врачей по направлениям, 2025", TITLE
    wg["A2"] = ("Вопрос из факторного анализа: приносят ли старшие грейды более высокий чек? Сравнивайте внутри направления: "
                "различия в целом по сети во многом объясняются составом направлений; внутри направления чек зависит ещё и от специальности, "
                "а врачей в ячейке мало, поэтому грейд и специальность перемешаны.")
    wg["A3"] = "На синтетических данных премии за грейд в чеке нет; в реальной клинике картина может быть другой."
    wg["A2"].font = wg["A3"].font = NOTE
    hg = ["Направление", "Грейд", "Врачей", "Визитов", "Выручка, тыс. ₽", "Чек, ₽", "Чек к среднему по направлению",
          "Доля оплаты с взносами", "Маржа после ФОТ, тыс. ₽", "Маржа, % выручки"]

    def grade_table(top, title, rows, by_dir):
        wg.cell(top, 1, title).font = BOLD
        for i, t in enumerate(hg, 1):
            wg.cell(top + 1, i, t)
        head(wg, top + 1, 1, len(hg))
        wg.row_dimensions[top + 1].height = 44
        rr = top + 2
        for dcode, g in rows:
            wg.cell(rr, 1, DIRS[dcode] if dcode != "*" else "Все направления").font = BLACK
            wg.cell(rr, 2, g).font = BLACK
            wg.cell(rr, 12, dcode).font = Font(name=F, size=8, color="808080")
            crit = f"{D('F')},$B{rr}" + (f",{D('D')},$L{rr}" if dcode != "*" else "")
            wg.cell(rr, 3, f"=COUNTIFS({doc_sheet_rng('E')},$B{rr}" + (f",{doc_sheet_rng('R')},$L{rr})" if dcode != "*" else ")"))
            wg.cell(rr, 4, f"=SUMIFS({D('G')},{crit})")
            wg.cell(rr, 5, f"=SUMIFS({D('H')},{crit})")
            wg.cell(rr, 6, f"=IF(D{rr}=0,0,E{rr}/D{rr})")
            avg = (f"SUMIFS({D('H')},{D('D')},$L{rr})/SUMIFS({D('G')},{D('D')},$L{rr})" if dcode != "*"
                   else f"SUM({D('H')})/SUM({D('G')})")
            wg.cell(rr, 7, f"=IF(F{rr}=0,0,F{rr}/({avg}))")
            wg.cell(rr, 8, f"=IF(E{rr}=0,0,SUMIFS({D('J')},{crit})/E{rr})")
            wg.cell(rr, 9, f"=SUMIFS({D('K')},{crit})")
            wg.cell(rr, 10, f"=IF(E{rr}=0,0,I{rr}/E{rr})")
            for c in range(3, 11):
                wg.cell(rr, c).font = BLACK
            wg.cell(rr, 3).number_format = wg.cell(rr, 4).number_format = "#,##0"
            wg.cell(rr, 5).number_format = wg.cell(rr, 9).number_format = NUM
            wg.cell(rr, 6).number_format = RUB
            for c in (7, 8, 10):
                wg.cell(rr, c).number_format = PCT
            rr += 1
        return rr

    combos = (doc.groupby(["direction_code", "grade"]).size().reset_index()
              .assign(o=lambda d: d.direction_code.map(order), g=lambda d: d.grade.map(gorder)).sort_values(["o", "g"]))
    nxt = grade_table(5, "По направлениям и грейдам", list(zip(combos.direction_code, combos.grade)), True)
    grade_table(nxt + 2, "По грейдам в целом (смесь направлений: чек сравнивать нельзя)", [("*", g) for g in GRADES], False)
    for i, w in enumerate([26, 10, 9, 11, 14, 11, 14, 11, 15, 12], 1):
        wg.column_dimensions[col(i)].width = w
    wg.column_dimensions["L"].width = 7
    wg.freeze_panes = "C7"

    # ---------------------------------------------------------------- Пациенты
    wp = wb.create_sheet("Пациенты", 2)
    wp["A1"], wp["A1"].font = "Повторные визиты и LTV-90 новых пациентов, 2025", TITLE
    wp["A2"] = ("Значения рассчитаны скриптом (не формулы). Когорта: пациенты с первым визитом в 2025 году до 02.10.2025. "
                "LTV-90: выручка пациента за 90 дней от первого визита. «Вернулись»: был ещё визит за 90 дней.")
    wp["A3"] = ("Распределение выручки на пациента скошено вправо: среднее заметно выше медианы, поэтому приведены оба. "
                "Данные есть только за 2025 год, длиннее 90 дней LTV не считается.")
    wp["A2"].font = wp["A3"].font = NOTE

    def seg(top, title, key, names):
        wp.cell(top, 1, title).font = BOLD
        for i, t in enumerate(["Сегмент", "Пациентов", "Вернулись за 90 дней", "LTV-90, медиана, ₽", "LTV-90, среднее, ₽"], 1):
            wp.cell(top + 1, i, t)
        head(wp, top + 1, 1, 5)
        wp.row_dimensions[top + 1].height = 30
        rr = top + 2
        groups = [(n, co[co[key] == k]) for k, n in names] + [("Все новые пациенты", co)]
        for nm, g in groups:
            vals = [nm, len(g), float(g.came_back.astype(float).mean()), float(g.ltv.median()), float(g.ltv.mean())]
            for i, v in enumerate(vals, 1):
                c = wp.cell(rr, i, v)
                c.font = BOLD if nm.startswith("Все") else BLACK
                c.number_format = {2: "#,##0", 3: PCT, 4: RUB, 5: RUB}.get(i, "General")
            if nm.startswith("Все"):
                for i in range(1, 6):
                    wp.cell(rr, i).border = TOP
            rr += 1
        return rr

    srcs = [(s, s) for s in co.groupby("source").size().sort_values(ascending=False).index]
    nxt = seg(5, "По источнику привлечения", "source", srcs)
    seg(nxt + 2, "По направлению первого визита", "first_dir", [(k, v) for k, v in DIRS.items()])
    for i, w in enumerate([30, 12, 16, 18, 18], 1):
        wp.column_dimensions[col(i)].width = w

    # ---------------------------------------------------------------- Сезонность
    wz = wb.create_sheet("Сезонность", 3)
    wz["A1"], wz["A1"].font = "Сезонность выручки 2025: индекс месяца к среднему, % (100 = средний месяц)", TITLE
    wz["A2"], wz["A2"].font = "Клиника (фильтр):", BOLD
    wz["B2"], wz["B2"].fill, wz["B2"].font = "Все", INPUT_FILL, BLUE
    wz["C2"] = '=IF($B$2="Все","*",$B$2)'
    wz["C2"].font = Font(name=F, size=8, color="808080")
    dv = DataValidation(type="list", formula1='"Все,C01,C02,C03"', allow_blank=False)
    wz.add_data_validation(dv)
    dv.add("B2")
    wz["D2"] = "← Все, C01, C02 или C03"
    wz["D2"].font = NOTE
    wz["A3"] = ("Индекс одного года не разделяет сезонность, рост и разовые события: выручка в декабре втрое выше январской, и план 2025 "
                "повторяет ту же форму (таблица «Факт к плану»). Месяцы с известными событиями отмечены ●. Прогноз не строится.")
    wz["A3"].font = NOTE
    periods = [f"{YEAR}-{m:02d}" for m in range(1, 13)]
    keys = list(DIRS) + ["*"]
    names = list(DIRS.values()) + ["Всего"]

    def month_head(row, title):
        wz.cell(row, 1, title)
        for i, p in enumerate(periods):
            wz.cell(row, 2 + i, p)
        head(wz, row, 1, 14)

    t1 = 5
    month_head(t1, "Выручка, тыс. ₽")
    wz.cell(t1, 14, "Среднее в месяц")
    for j, (k, nm) in enumerate(zip(keys, names)):
        rr = t1 + 1 + j
        wz.cell(rr, 1, nm).font = BOLD if k == "*" else BLACK
        wz.cell(rr, 16, k).font = Font(name=F, size=8, color="808080")
        for i in range(12):
            L = col(2 + i)
            wz.cell(rr, 2 + i, f"=SUMIFS({D('H')},{D('A')},{L}${t1},{D('D')},$P{rr},{D('E')},$C$2)")
            wz.cell(rr, 2 + i).number_format = NUM
            wz.cell(rr, 2 + i).font = BOLD if k == "*" else BLACK
        wz.cell(rr, 14, f"=AVERAGE(B{rr}:M{rr})").number_format = NUM
        wz.cell(rr, 14).font = BOLD if k == "*" else BLACK
    t2 = t1 + len(keys) + 3
    month_head(t2, "Индекс месяца к среднему, %")
    for j, (k, nm) in enumerate(zip(keys, names)):
        rr = t2 + 1 + j
        src = t1 + 1 + j
        wz.cell(rr, 1, nm).font = BOLD if k == "*" else BLACK
        wz.cell(rr, 16, k).font = Font(name=F, size=8, color="808080")
        for i in range(12):
            L = col(2 + i)
            wz.cell(rr, 2 + i, f"=IF($N{src}=0,0,{L}{src}/$N{src}*100)")
            wz.cell(rr, 2 + i).number_format = "0"
            wz.cell(rr, 2 + i).font = BOLD if k == "*" else BLACK
    wz.conditional_formatting.add(f"B{t2 + 1}:M{t2 + len(keys)}", ColorScaleRule(
        start_type="num", start_value=80, start_color="F8696B", mid_type="num", mid_value=100, mid_color="FFFFFF",
        end_type="num", end_value=115, end_color="63BE7B"))
    t3 = t2 + len(keys) + 3
    month_head(t3, "Известные разовые события (●)")
    t4 = t3 + len(DIRS) + 3
    month_head(t4, "План выручки, тыс. ₽ (вся сеть; значения из budget_revenue)")
    pl = plan.pivot(index="direction_code", columns="period", values="plan_rev").reindex(index=list(DIRS), columns=periods).fillna(0)
    for j, (k, nm) in enumerate(zip(keys, names)):
        rr = t4 + 1 + j
        wz.cell(rr, 1, nm).font = BOLD if k == "*" else BLACK
        wz.cell(rr, 16, k).font = Font(name=F, size=8, color="808080")
        for i, pnm in enumerate(periods):
            L = col(2 + i)
            if k == "*":
                wz.cell(rr, 2 + i, f"=SUM({L}{t4 + 1}:{L}{t4 + len(DIRS)})").font = BOLD
            else:
                wz.cell(rr, 2 + i, round(float(pl.loc[k, pnm]), 2)).font = BLUE
            wz.cell(rr, 2 + i).number_format = NUM
    t5 = t4 + len(keys) + 3
    month_head(t5, "Факт к плану, % (вся сеть, фильтр клиники не действует)")
    for j, (k, nm) in enumerate(zip(keys, names)):
        rr = t5 + 1 + j
        wz.cell(rr, 1, nm).font = BOLD if k == "*" else BLACK
        wz.cell(rr, 16, k).font = Font(name=F, size=8, color="808080")
        for i in range(12):
            L = col(2 + i)
            wz.cell(rr, 2 + i, f"=IF({L}{t4 + 1 + j}=0,0,SUMIFS({D('H')},{D('A')},{L}${t5},{D('D')},$P{rr})/{L}{t4 + 1 + j})")
            wz.cell(rr, 2 + i).number_format = "0%"
            wz.cell(rr, 2 + i).font = BOLD if k == "*" else BLACK
    wz.conditional_formatting.add(f"B{t5 + 1}:M{t5 + len(keys)}", ColorScaleRule(
        start_type="num", start_value=0.8, start_color="F8696B", mid_type="num", mid_value=0.95, mid_color="FFFFFF",
        end_type="num", end_value=1.05, end_color="63BE7B"))
    ev0 = t5 + len(keys) + 5          # первая строка таблицы событий
    ev_n = len(EVENTS) + 4            # запас строк под новые события
    EP, ED, EC = (f"$A${ev0}:$A${ev0 + ev_n - 1}", f"$B${ev0}:$B${ev0 + ev_n - 1}", f"$C${ev0}:$C${ev0 + ev_n - 1}")
    for j, (k, nm) in enumerate(zip(list(DIRS), list(DIRS.values()))):
        rr = t3 + 1 + j
        wz.cell(rr, 1, nm).font = BLACK
        wz.cell(rr, 16, k).font = Font(name=F, size=8, color="808080")
        for i in range(12):
            L = col(2 + i)
            f_all = f"COUNTIFS({EP},{L}${t3},{ED},$P{rr})"
            f_clin = f"COUNTIFS({EP},{L}${t3},{ED},$P{rr},{EC},\"Все\")+COUNTIFS({EP},{L}${t3},{ED},$P{rr},{EC},$B$2)"
            wz.cell(rr, 2 + i, f'=IF(IF($B$2="Все",{f_all},{f_clin})>0,"●","")')
            wz.cell(rr, 2 + i).font = RED
            wz.cell(rr, 2 + i).alignment = Alignment(horizontal="center")
    wz.cell(ev0 - 1, 1, "Период")
    wz.cell(ev0 - 1, 2, "Направление (код)")
    wz.cell(ev0 - 1, 3, "Клиника (Все или код)")
    wz.cell(ev0 - 1, 4, "Событие")
    head(wz, ev0 - 1, 1, 4)
    wz.cell(ev0 - 2, 1, "Таблица событий (вход; строки ниже можно добавлять)").font = BOLD
    for i in range(ev_n):
        row = EVENTS[i] if i < len(EVENTS) else ("", "", "", "")
        for c, v in enumerate(row, 1):
            cell = wz.cell(ev0 + i, c, v)
            cell.font, cell.fill = BLUE, INPUT_FILL
    wz.cell(ev0 + ev_n, 1, "Майские праздники: календарь. Простой лазера C03: сведения клиники. Оба события известны из внешнего источника, "
                           "не выведены из данных.").font = NOTE

    ch = LineChart()
    ch.title = "Индекс месяца к среднему, %"
    for j in range(len(keys)):
        ch.add_data(Reference(wz, min_col=1, max_col=13, min_row=t2 + 1 + j), from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(wz, min_col=2, max_col=13, min_row=t2))
    ch.y_axis.numFmt = "0"
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    ch.height, ch.width = 9, 24
    wz.add_chart(ch, f"A{ev0 + ev_n + 3}")
    wz.column_dimensions["A"].width = 30
    for c in range(2, 15):
        wz.column_dimensions[col(c)].width = 11
    wz.column_dimensions["D"].width = 11
    wz.column_dimensions["P"].width = 7
    wz.freeze_panes = "B6"

    # ---------------------------------------------------------------- Описание
    wi = wb.create_sheet("Описание", 0)
    wi.column_dimensions["A"].width = 125
    text = [
        ("Аналитика врачей, 2025", TITLE),
        ("Что здесь", BOLD),
        ("Врачи: выручка, визиты, чек и маржа после оплаты труда по каждому врачу и строка «Без врача». Грейды: чек, доля "
         "оплаты и маржа по направлениям. Пациенты: повторные визиты и LTV-90. Сезонность: индекс месяца к среднему по "
         "направлениям с отметкой известных событий.", BLACK),
        ("Как считаем", BOLD),
        ("Выручка и визиты: завершённые визиты ERP после исправлений сверки. Материалы и лаборатории: по нормативу услуг "
         "(количество × норма единицы). Оплата врача: начислено и взносы из payroll_1c. Маржа после ФОТ = выручка − "
         "материалы и лаборатории − оплата врача с взносами.", BLACK),
        ("Врачи сравниваются внутри направления и специальности. Чек к среднему по направлению показывает, где врач выше или ниже "
         "типичного для направления. Доля оплаты показана с взносами (36% по справочнику ×1,22 ≈ 43,9%).", BLACK),
        ("Ограничения", BOLD),
        ("1. Взносы: ставка 22% плоская, допущение демо (без предельной базы). Оплата врача берётся из payroll_1c, поэтому в ней "
         "уже заложена эта ставка.", BLACK),
        ("2. Два визита без врача (120 700 ₽) показаны отдельной строкой. Оплата врачам за них в 1С была начислена до "
         "исправления, поэтому у остальных врачей оплата чуть выше выручки ERP.", BLACK),
        ("3. Данные только за 2025 год: LTV считается за 90 дней, повторные визиты тоже за 90 дней; пациенты до 2025 года "
         "известны только датой первого визита.", BLACK),
        ("4. Индекс одного года не отделяет сезонность от роста и разовых событий: выручка в декабре втрое выше январской, и "
         "план 2025 повторяет ту же форму. Разовые события (майские праздники, простой лазера C03) отмечены и видны на таблице "
         "«Факт к плану». Прогноз на 2026 не строится: для этого нужна история минимум за 2-3 года.", BLACK),
        ("5. У направления 1-8 врачей: показатели конкретного врача по одному году нельзя считать оценкой его работы.", BLACK),
        ("6. На синтетических данных премии за грейд в чеке нет; в реальной клинике картина может отличаться.", BLACK),
        ("Цвета: синий шрифт: входные числа, зелёный: ссылка на лист «Данные», чёрный: расчёты, жёлтая заливка: фильтры и таблица событий.", NOTE),
    ]
    for i, (t, f) in enumerate(text, 1):
        c = wi.cell(i, 1, t)
        c.font, c.alignment = f, Alignment(wrap_text=True, vertical="top")
    return wb, nodoc_r


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    data, doctors, ret, co, plan = prepare()
    wb, nodoc_r = build(data, doctors, ret, co, plan)
    path = OUT / "DoctorAnalytics_2025.xlsx"
    wb.save(path)
    recalculate(path)
    print(f"Сохранено: {path}  (строк данных: {len(data)})")
    print(f"Визитов {data.visits.sum():,}, выручка {data.revenue.sum():,.0f} ₽, без врача: "
          f"{data[data.doctor_id == NODOC].visits.sum()} визитов на {data[data.doctor_id == NODOC].revenue.sum():,.0f} ₽")
    print(f"Оплата врачей с взносами {data.pay_total.sum():,.0f} ₽; маржа после ФОТ "
          f"{(data.revenue - data.norm_cost - data.pay_total).sum():,.0f} ₽")
    print(f"Когорта повторных визитов: {int(ret.cohort_n.sum()):,}; когорта LTV-90: {len(co):,}; "
          f"LTV-90 медиана {co.ltv.median():,.0f} / среднее {co.ltv.mean():,.0f} ₽; вернулись {co.came_back.astype(float).mean():.1%}")


if __name__ == "__main__":
    main()

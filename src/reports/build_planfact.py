"""
Шаг 10: план-факт по всему P&L за 2025 год в Excel.

Вход:  data/corrected/ - проводки 1С (факт), budget_revenue.csv и budget_opex.csv (план),
                         visits.csv и visit_items.csv (объёмы услуг)
Выход: data/reports/PlanFact_2025.xlsx

Факт берётся тем же кодом, что и в PnL_2025.xlsx (build_pnl.prepare), поэтому итоги совпадают.
Книга живая: листы считаются формулами SUMIFS по листам «Данные» и «Объёмы».

Листы:
  Описание           - источники, правила знаков, ограничения
  План-факт          - годовой план-факт по всем статьям P&L + влияние на прибыль по месяцам (фильтр по клинике)
  По клиникам        - план, факт и влияние на прибыль по каждой клинике и сети
  По направлениям    - объём, выручка, средний чек и прямая маржа (после материалов и лабораторий)
  Тепловая карта     - отклонение выручки и объёма услуг в % по направлениям и месяцам (фильтр по клинике)
  Топ отклонений     - самые крупные отклонения выручки по клиника x направление x месяц и по статьям
  Данные, Объёмы     - источники для формул

Знак «Влияние на прибыль»: плюс = выгодно для прибыли (выручка выше плана или расход ниже плана),
минус = невыгодно. «Отклонение» = факт минус план в исходном знаке.

Запуск:  python src/reports/build_planfact.py
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.recalc import recalculate  # noqa: E402
import build_pnl  # noqa: E402  (факт: ту же подготовку использует PnL_2025.xlsx)

ROOT = Path(__file__).resolve().parents[2]
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "reports"
YEAR = 2025
CLINICS = ["C01", "C02", "C03"]
DIRS = {"DENT": "Стоматология", "FAM": "Семейная медицина", "COSM": "Косметология",
        "PLAS": "Пластическая хирургия", "LASER": "Лазерная коррекция зрения"}
MAT, LAB = "Материалы", "Лаборатории и внешние услуги"
DIRECT = ["Выручка", MAT, LAB]            # статьи, которые есть в плане по направлениям
BLOCKS = {"cost_of_sales": "Себестоимость", "management": "Управленческие", "commercial": "Коммерческие"}

COST_ARTICLES = ["Оплата труда врачей", "Оплата труда медперсонала", "Страховые взносы (себестоимость)",
                 MAT, LAB, "Амортизация"]
MGMT_ARTICLES = ["Оплата труда административного персонала", "Страховые взносы (администрация)", "Аренда",
                 "Коммунальные услуги", "Комиссия эквайринга"]
COMM_ARTICLES = ["Маркетинг"]

F = "Arial"
BLUE = Font(name=F, size=10, color="0000FF")
BLACK = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
TITLE = Font(name=F, size=14, bold=True)
HDR_FONT = Font(name=F, size=10, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SEC_FILL = PatternFill("solid", fgColor="D9E2F3")
INPUT_FILL = PatternFill("solid", fgColor="FFFF99")
TOP = Border(top=Side(style="thin"))
NUM = "#,##0,;(#,##0,);-"
PCT = "0.0%;(0.0%);-"
PP = '0.0" п.п.";(0.0" п.п.");-'
RED = Font(name=F, size=10, color="C00000")
GREEN = Font(name=F, size=10, color="00803C")


# ------------------------------------------------------------------ данные
def prepare():
    fact, _ = build_pnl.prepare()
    fact = fact.copy()
    fact["direction_code"] = fact.apply(lambda r: r.direction_code if r.article in DIRECT else "ОБЩ", axis=1)
    fact = (fact.groupby(["period", "clinic_id", "direction_code", "article"], as_index=False)
            .amount.sum().rename(columns={"amount": "fact"}))

    br = pd.read_csv(CORR / "budget_revenue.csv")
    br["mat"] = br.plan_qty * br.plan_unit_material
    br["lab"] = br.plan_qty * br.plan_unit_external
    g = br.groupby(["period", "clinic_id", "direction_code"], as_index=False)[
        ["plan_revenue", "mat", "lab"]].sum()
    plan = pd.concat([
        g.rename(columns={"plan_revenue": "plan"})[["period", "clinic_id", "direction_code", "plan"]].assign(article="Выручка"),
        g.rename(columns={"mat": "plan"})[["period", "clinic_id", "direction_code", "plan"]].assign(article=MAT),
        g.rename(columns={"lab": "plan"})[["period", "clinic_id", "direction_code", "plan"]].assign(article=LAB),
    ])
    bo = pd.read_csv(CORR / "budget_opex.csv").rename(columns={"plan_amount": "plan"})
    bo["direction_code"] = "ОБЩ"
    plan = pd.concat([plan, bo[["period", "clinic_id", "direction_code", "article", "plan"]]])
    plan = plan.groupby(["period", "clinic_id", "direction_code", "article"], as_index=False).plan.sum()

    data = plan.merge(fact, on=["period", "clinic_id", "direction_code", "article"], how="outer").fillna(0.0)
    data["month"] = data.period.str[5:7].astype(int)
    data = data[["period", "month", "clinic_id", "direction_code", "article", "plan", "fact"]]
    data[["plan", "fact"]] = data[["plan", "fact"]].round(2)
    data = data.sort_values(["period", "clinic_id", "direction_code", "article"]).reset_index(drop=True)

    # объёмы услуг: план и факт, штук
    visits = pd.read_csv(CORR / "visits.csv")
    items = pd.read_csv(CORR / "visit_items.csv")
    services = pd.read_csv(CORR / "services.csv")[["service_id", "direction_code"]]
    im = items.merge(visits[["visit_id", "visit_date", "clinic_id"]], on="visit_id").merge(services, on="service_id")
    im["period"] = im.visit_date.str[:7]
    fq = im.groupby(["period", "clinic_id", "direction_code"], as_index=False).qty.sum().rename(columns={"qty": "fact_qty"})
    pq = br.groupby(["period", "clinic_id", "direction_code"], as_index=False).plan_qty.sum()
    qty = pq.merge(fq, on=["period", "clinic_id", "direction_code"], how="outer").fillna(0)
    qty["month"] = qty.period.str[5:7].astype(int)
    qty = qty[["period", "month", "clinic_id", "direction_code", "plan_qty", "fact_qty"]]
    qty = qty.sort_values(["period", "clinic_id", "direction_code"]).reset_index(drop=True)
    return data, qty


# ------------------------------------------------------------------ оформление
def head(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build(data, qty):
    wb = Workbook()

    # ---------- Данные ----------
    wd = wb.active
    wd.title = "Данные"
    wd.append(list(data.columns))
    for r in data.itertuples(index=False):
        wd.append(list(r))
    nd = wd.max_row
    head(wd, 1, 1, 7)
    for row in wd.iter_rows(min_row=2, max_row=nd):
        for c in row:
            c.font = BLUE if c.column >= 6 else BLACK
            if c.column >= 6:
                c.number_format = "#,##0.00"
    for i, w in enumerate([10, 8, 10, 14, 40, 16, 16], 1):
        wd.column_dimensions[col(i)].width = w
    wd.freeze_panes = "A2"
    D = lambda c: f"'Данные'!${c}$2:${c}${nd}"          # noqa: E731

    # ---------- Объёмы ----------
    wq = wb.create_sheet("Объёмы")
    wq.append(list(qty.columns))
    for r in qty.itertuples(index=False):
        wq.append(list(r))
    nq = wq.max_row
    head(wq, 1, 1, 6)
    for row in wq.iter_rows(min_row=2, max_row=nq):
        for c in row:
            c.font = BLUE if c.column >= 5 else BLACK
    for i, w in enumerate([10, 8, 10, 14, 12, 12], 1):
        wq.column_dimensions[col(i)].width = w
    wq.freeze_panes = "A2"
    Q = lambda c: f"'Объёмы'!${c}$2:${c}${nq}"          # noqa: E731

    periods = [f"{YEAR}-{m:02d}" for m in range(1, 13)]

    # ======================= План-факт =======================
    ws = wb.create_sheet("План-факт", 0)
    ws["A1"], ws["A1"].font = "План-факт по P&L сети клиник, 2025 (тыс. ₽)", TITLE
    ws["A2"], ws["A2"].font = "Клиника (фильтр):", BOLD
    ws["B2"], ws["B2"].fill, ws["B2"].font = "Все", INPUT_FILL, BLUE
    ws["C2"] = '=IF($B$2="Все","*",$B$2)'
    ws["C2"].font = Font(name=F, size=8, color="808080")
    ws["D2"] = "← «Все», C01, C02 или C03"
    ws["D2"].font = Font(name=F, size=9, italic=True, color="808080")
    dv = DataValidation(type="list", formula1='"Все,C01,C02,C03"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B2")
    ws["A3"] = "Влияние на прибыль: плюс выгодно (выручка выше плана, расход ниже плана), минус невыгодно."
    ws["A3"].font = Font(name=F, size=9, italic=True)

    for i, h in enumerate(["Статья", "План", "Факт", "Отклонение (факт − план)", "Влияние на прибыль",
                           "Откл., %", "% выручки, план", "% выручки, факт"], 1):
        ws.cell(5, i, h)
    head(ws, 5, 1, 8)
    ws.row_dimensions[5].height = 30
    ws.column_dimensions["A"].width = 44
    for c in range(2, 15):
        ws.column_dimensions[col(c)].width = 13
    ws.freeze_panes = "B6"

    rows = {}
    r = 6

    def line(label, key=None, article=None, sign=-1):
        """Строка из данных: sign=+1 для выручки, -1 для расходов."""
        nonlocal r
        ws.cell(r, 1, label).font = BLACK
        for c, fld in ((2, "F"), (3, "G")):
            ws.cell(r, c, f'=SUMIFS({D(fld)},{D("E")},"{article}",{D("C")},$C$2)')
        ws.cell(r, 4, f"=C{r}-B{r}")
        ws.cell(r, 5, f"=D{r}" if sign > 0 else f"=-D{r}")
        ws.cell(r, 6, f"=IF(B{r}=0,0,D{r}/B{r})")
        rows[key or article] = r
        r += 1

    def total(label, key, formula_b, bold=True, fill=None):
        """Итоговая строка: formula_b(L) возвращает формулу для столбцов B, C и E."""
        nonlocal r
        ws.cell(r, 1, label).font = BOLD
        for c in (2, 3, 5):
            ws.cell(r, c, formula_b(col(c)))
        ws.cell(r, 4, f"=C{r}-B{r}")
        ws.cell(r, 6, f"=IF(B{r}=0,0,D{r}/B{r})")
        for c in range(1, 9):
            ws.cell(r, c).border = TOP
            if fill:
                ws.cell(r, c).fill = fill
        rows[key] = r
        r += 1

    def section(label):
        nonlocal r
        for c in range(1, 9):
            ws.cell(r, c).fill = SEC_FILL
        ws.cell(r, 1, label).font = BOLD
        r += 1

    line("Выручка", article="Выручка", sign=1)
    section("Себестоимость услуг")
    for a in COST_ARTICLES:
        line(a, article=a)
    c1, c2 = rows[COST_ARTICLES[0]], rows[COST_ARTICLES[-1]]
    total("Итого себестоимость", "cost", lambda L: f"=SUM({L}{c1}:{L}{c2})")
    # E для итога: сумма влияний
    ws.cell(rows["cost"], 5, f"=SUM(E{c1}:E{c2})")
    total("Валовая прибыль", "gross", lambda L: f"={L}{rows['Выручка']}-{L}{rows['cost']}")
    ws.cell(rows["gross"], 5, f"=E{rows['Выручка']}+E{rows['cost']}")
    section("Управленческие расходы")
    for a in MGMT_ARTICLES:
        line(a, article=a)
    m1, m2 = rows[MGMT_ARTICLES[0]], rows[MGMT_ARTICLES[-1]]
    total("Итого управленческие", "mgmt", lambda L: f"=SUM({L}{m1}:{L}{m2})")
    ws.cell(rows["mgmt"], 5, f"=SUM(E{m1}:E{m2})")
    section("Коммерческие расходы")
    line("Маркетинг", article="Маркетинг")
    total("Прибыль до налога", "profit", lambda L: f"={L}{rows['gross']}-{L}{rows['mgmt']}-{L}{rows['Маркетинг']}",
          fill=SEC_FILL)
    ws.cell(rows["profit"], 5, f"=E{rows['gross']}+E{rows['mgmt']}+E{rows['Маркетинг']}")
    total("EBITDA", "ebitda", lambda L: f"={L}{rows['profit']}+{L}{rows['Амортизация']}")
    ws.cell(rows["ebitda"], 5, f"=E{rows['profit']}+E{rows['Амортизация']}")
    last_amount_row = r - 1
    rev = rows["Выручка"]
    for rr in range(6, last_amount_row + 1):
        if ws.cell(rr, 2).value is None:
            continue
        ws.cell(rr, 7, f"=IF($B${rev}=0,0,B{rr}/$B${rev})")
        ws.cell(rr, 8, f"=IF($C${rev}=0,0,C{rr}/$C${rev})")
        for c in range(2, 6):
            ws.cell(rr, c).number_format = NUM
        for c in (6, 7, 8):
            ws.cell(rr, c).number_format = PCT
        bold = rr in (rows["cost"], rows["gross"], rows["mgmt"], rows["profit"], rows["ebitda"])
        for c in range(2, 9):
            ws.cell(rr, c).font = BOLD if bold else BLACK
    # маржи
    for label, key, num in [("Валовая маржа, %", "gm", "gross"), ("Маржа прибыли, %", "pm", "profit"),
                            ("EBITDA, %", "em", "ebitda")]:
        ws.cell(r, 1, label).font = BOLD
        ws.cell(r, 2, f"=IF(B{rev}=0,0,B{rows[num]}/B{rev})")
        ws.cell(r, 3, f"=IF(C{rev}=0,0,C{rows[num]}/C{rev})")
        ws.cell(r, 4, f"=(C{r}-B{r})*100")
        for c in (2, 3):
            ws.cell(r, c).number_format = PCT
            ws.cell(r, c).font = BOLD
        ws.cell(r, 4).number_format = PP
        ws.cell(r, 4).font = BOLD
        rows[key] = r
        r += 1
    # условное форматирование влияния
    rng = f"E6:E{last_amount_row}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=RED))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=GREEN))

    # --- помесячно ---
    r += 1
    ws.cell(r, 1, "Выручка: план, факт и отклонение по месяцам").font = BOLD
    r += 1
    mh = r
    ws.cell(r, 1, "Показатель")
    for i, p in enumerate(periods):
        ws.cell(r, 2 + i, p)
    ws.cell(r, 14, f"Итого {YEAR}")
    head(ws, r, 1, 14)
    r += 1
    m_plan, m_fact, m_pct = r, r + 1, r + 2
    ws.cell(m_plan, 1, "Выручка, план")
    ws.cell(m_fact, 1, "Выручка, факт")
    ws.cell(m_pct, 1, "Отклонение, %").font = BOLD
    for i in range(12):
        L = col(2 + i)
        for rr, fld in ((m_plan, "F"), (m_fact, "G")):
            ws.cell(rr, 2 + i, f'=SUMIFS({D(fld)},{D("A")},{L}${mh},{D("E")},"Выручка",{D("C")},$C$2)')
        ws.cell(m_pct, 2 + i, f"=IF({L}{m_plan}=0,0,{L}{m_fact}/{L}{m_plan}-1)")
    ws.cell(m_plan, 14, f"=SUM(B{m_plan}:M{m_plan})")
    ws.cell(m_fact, 14, f"=SUM(B{m_fact}:M{m_fact})")
    ws.cell(m_pct, 14, f"=IF(N{m_plan}=0,0,N{m_fact}/N{m_plan}-1)")
    for rr in (m_plan, m_fact):
        for c in range(2, 15):
            ws.cell(rr, c).number_format = NUM
            ws.cell(rr, c).font = BLACK
    for c in range(2, 15):
        ws.cell(m_pct, c).number_format = PCT
        ws.cell(m_pct, c).font = BOLD
    ws.conditional_formatting.add(f"B{m_pct}:N{m_pct}", ColorScaleRule(
        start_type="num", start_value=-0.12, start_color="F8696B",
        mid_type="num", mid_value=0, mid_color="FFFFFF", end_type="num", end_value=0.05, end_color="63BE7B"))
    r = m_pct + 2
    ws.cell(r, 1, "Влияние на прибыль по месяцам (тыс. ₽): плюс выгодно, минус невыгодно").font = BOLD
    r += 1
    ih = r
    ws.cell(r, 1, "Статья")
    for i, p in enumerate(periods):
        ws.cell(r, 2 + i, p)
    ws.cell(r, 14, f"Итого {YEAR}")
    head(ws, r, 1, 14)
    r += 1
    irows = {}

    def i_line(label, article, sign):
        nonlocal r
        ws.cell(r, 1, label).font = BLACK
        for i in range(12):
            L = col(2 + i)
            core = (f'SUMIFS({D("G")},{D("A")},{L}${ih},{D("E")},"{article}",{D("C")},$C$2)'
                    f'-SUMIFS({D("F")},{D("A")},{L}${ih},{D("E")},"{article}",{D("C")},$C$2)')
            ws.cell(r, 2 + i, f"={core}" if sign > 0 else f"=-({core})")
        ws.cell(r, 14, f"=SUM(B{r}:M{r})")
        irows[article] = r
        r += 1

    def i_total(label, key, parts, bold_fill=None):
        nonlocal r
        ws.cell(r, 1, label).font = BOLD
        for c in range(2, 15):
            L = col(c)
            ws.cell(r, c, "=" + "+".join(f"{L}{p}" for p in parts))
            ws.cell(r, c).border = TOP
            if bold_fill:
                ws.cell(r, c).fill = bold_fill
        ws.cell(r, 1).border = TOP
        irows[key] = r
        r += 1

    i_line("Выручка", "Выручка", +1)
    for a in COST_ARTICLES:
        i_line(a, a, -1)
    i_total("Итого себестоимость", "cost", [irows[a] for a in COST_ARTICLES])
    i_total("Валовая прибыль", "gross", [irows["Выручка"], irows["cost"]])
    for a in MGMT_ARTICLES:
        i_line(a, a, -1)
    i_total("Итого управленческие", "mgmt", [irows[a] for a in MGMT_ARTICLES])
    i_line("Маркетинг", "Маркетинг", -1)
    i_total("Прибыль до налога", "profit", [irows["gross"], irows["mgmt"], irows["Маркетинг"]], SEC_FILL)
    for rr in range(ih + 1, r):
        for c in range(2, 15):
            ws.cell(rr, c).number_format = NUM
            ws.cell(rr, c).font = BOLD if rr in (irows["cost"], irows["gross"], irows["mgmt"], irows["profit"]) else BLACK
    ws.conditional_formatting.add(f"B{ih + 1}:N{r - 1}", CellIsRule(operator="lessThan", formula=["0"], font=RED))
    ws.conditional_formatting.add(f"B{ih + 1}:N{r - 1}", CellIsRule(operator="greaterThan", formula=["0"], font=GREEN))

    # ======================= По клиникам =======================
    wk = wb.create_sheet("По клиникам", 1)
    wk["A1"], wk["A1"].font = "План-факт по клиникам, 2025 (тыс. ₽)", TITLE
    wk["A2"] = "Влияние на прибыль: плюс выгодно, минус невыгодно."
    wk["A2"].font = Font(name=F, size=9, italic=True)
    groups = [(c, f'"{c}"') for c in CLINICS] + [("Сеть", '"*"')]
    wk.cell(4, 1, "Статья")
    for gi, (name, _) in enumerate(groups):
        for j, h in enumerate(["План", "Факт", "Влияние"]):
            wk.cell(4, 2 + gi * 3 + j, f"{name}: {h}")
    head(wk, 4, 1, 13)
    wk.row_dimensions[4].height = 30
    wk.column_dimensions["A"].width = 44
    for c in range(2, 14):
        wk.column_dimensions[col(c)].width = 13
    wk.freeze_panes = "B5"
    kr = 5
    krow = {}

    def k_line(label, article, sign):
        nonlocal kr
        wk.cell(kr, 1, label).font = BLACK
        for gi, (_, crit) in enumerate(groups):
            b = 2 + gi * 3
            wk.cell(kr, b, f'=SUMIFS({D("F")},{D("E")},"{article}",{D("C")},{crit})')
            wk.cell(kr, b + 1, f'=SUMIFS({D("G")},{D("E")},"{article}",{D("C")},{crit})')
            wk.cell(kr, b + 2, f"={col(b + 1)}{kr}-{col(b)}{kr}" if sign > 0 else f"={col(b)}{kr}-{col(b + 1)}{kr}")
        krow[article] = kr
        kr += 1

    def k_total(label, key, parts_pf, parts_inf, fill=None):
        """parts_pf: список (знак, строка) для плана и факта; parts_inf: строки для влияния (сумма)."""
        nonlocal kr
        wk.cell(kr, 1, label).font = BOLD
        for gi in range(4):
            b = 2 + gi * 3
            for off in (0, 1):
                wk.cell(kr, b + off, "=" + "".join(f"{'+' if s > 0 else '-'}{col(b + off)}{rr}" for s, rr in parts_pf))
            wk.cell(kr, b + 2, "=" + "+".join(f"{col(b + 2)}{rr}" for rr in parts_inf))
        for c in range(1, 14):
            wk.cell(kr, c).border = TOP
            if fill:
                wk.cell(kr, c).fill = fill
        krow[key] = kr
        kr += 1

    k_line("Выручка", "Выручка", +1)
    for a in COST_ARTICLES:
        k_line(a, a, -1)
    cs = [krow[a] for a in COST_ARTICLES]
    k_total("Итого себестоимость", "cost", [(1, x) for x in cs], cs)
    k_total("Валовая прибыль", "gross", [(1, krow["Выручка"]), (-1, krow["cost"])], [krow["Выручка"], krow["cost"]])
    for a in MGMT_ARTICLES:
        k_line(a, a, -1)
    ms = [krow[a] for a in MGMT_ARTICLES]
    k_total("Итого управленческие", "mgmt", [(1, x) for x in ms], ms)
    k_line("Маркетинг", "Маркетинг", -1)
    k_total("Прибыль до налога", "profit",
            [(1, krow["gross"]), (-1, krow["mgmt"]), (-1, krow["Маркетинг"])],
            [krow["gross"], krow["mgmt"], krow["Маркетинг"]], SEC_FILL)
    wk.cell(kr, 1, "Валовая маржа, %").font = BOLD
    wk.cell(kr + 1, 1, "Маржа прибыли, %").font = BOLD
    for gi in range(4):
        b = 2 + gi * 3
        for off in (0, 1):
            L = col(b + off)
            wk.cell(kr, b + off, f"=IF({L}{krow['Выручка']}=0,0,{L}{krow['gross']}/{L}{krow['Выручка']})")
            wk.cell(kr + 1, b + off, f"=IF({L}{krow['Выручка']}=0,0,{L}{krow['profit']}/{L}{krow['Выручка']})")
        for rr in (kr, kr + 1):
            wk.cell(rr, b + 2, f"=({col(b + 1)}{rr}-{col(b)}{rr})*100")
            wk.cell(rr, b + 2).number_format = PP
            for off in (0, 1):
                wk.cell(rr, b + off).number_format = PCT
    last_k = kr + 1
    for rr in range(5, kr):
        bold = rr in (krow["cost"], krow["gross"], krow["mgmt"], krow["profit"])
        for c in range(2, 14):
            wk.cell(rr, c).number_format = NUM
            wk.cell(rr, c).font = BOLD if bold else BLACK
    for rr in (kr, kr + 1):
        for c in range(2, 14):
            wk.cell(rr, c).font = BOLD
    for gi in range(4):
        L = col(4 + gi * 3)
        rng = f"{L}5:{L}{kr - 1}"
        wk.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=RED))
        wk.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=GREEN))

    # ======================= По направлениям =======================
    wn = wb.create_sheet("По направлениям", 2)
    wn["A1"], wn["A1"].font = "План-факт по направлениям, 2025", TITLE
    wn["A2"], wn["A2"].font = "Клиника (фильтр):", BOLD
    wn["B2"], wn["B2"].fill, wn["B2"].font = "Все", INPUT_FILL, BLUE
    wn["C2"] = '=IF($B$2="Все","*",$B$2)'
    wn["C2"].font = Font(name=F, size=8, color="808080")
    dv2 = DataValidation(type="list", formula1='"Все,C01,C02,C03"', allow_blank=False)
    wn.add_data_validation(dv2)
    dv2.add("B2")
    wn["A3"] = ("Косвенные расходы (врачи, персонал, аренда и др.) в плане не разнесены по направлениям, поэтому "
                "прямая маржа = выручка − материалы − лаборатории. Выручка и затраты в тыс. ₽, объём в штуках услуг.")
    wn["A3"].font = Font(name=F, size=9, italic=True)
    hdrs = ["Направление", "Объём, план", "Объём, факт", "Объём, откл. %", "Выручка, план", "Выручка, факт",
            "Выручка, откл.", "Выручка, откл. %", "Ср. чек, план, ₽", "Ср. чек, факт, ₽", "Ср. чек, откл. %",
            "Материалы + лаборатории, план", "Материалы + лаборатории, факт", "Прямая маржа, план %",
            "Прямая маржа, факт %", "Прямая маржа, откл."]
    for i, h in enumerate(hdrs, 1):
        wn.cell(5, i, h)
    head(wn, 5, 1, len(hdrs))
    wn.row_dimensions[5].height = 45
    wn.column_dimensions["A"].width = 30
    for c in range(2, len(hdrs) + 1):
        wn.column_dimensions[col(c)].width = 13
    wn.freeze_panes = "B6"
    nr = 6
    for code, name in list(DIRS.items()) + [("*", "Итого")]:
        bold = code == "*"
        wn.cell(nr, 1, name).font = BOLD if bold else BLACK
        crit = '"*"' if bold else f'"{code}"'
        s = lambda fld, art: (f'SUMIFS({D(fld)},{D("E")},"{art}",{D("D")},{crit},{D("C")},$C$2)')   # noqa: E731
        qs = lambda fld: f'SUMIFS({Q(fld)},{Q("D")},{crit},{Q("C")},$C$2)'                         # noqa: E731
        wn.cell(nr, 2, f"={qs('E')}")
        wn.cell(nr, 3, f"={qs('F')}")
        wn.cell(nr, 4, f"=IF(B{nr}=0,0,C{nr}/B{nr}-1)")
        wn.cell(nr, 5, f"={s('F', 'Выручка')}")
        wn.cell(nr, 6, f"={s('G', 'Выручка')}")
        wn.cell(nr, 7, f"=F{nr}-E{nr}")
        wn.cell(nr, 8, f"=IF(E{nr}=0,0,F{nr}/E{nr}-1)")
        wn.cell(nr, 9, f"=IF(B{nr}=0,0,E{nr}/B{nr})")
        wn.cell(nr, 10, f"=IF(C{nr}=0,0,F{nr}/C{nr})")
        wn.cell(nr, 11, f"=IF(I{nr}=0,0,J{nr}/I{nr}-1)")
        wn.cell(nr, 12, f"={s('F', MAT)}+{s('F', LAB)}")
        wn.cell(nr, 13, f"={s('G', MAT)}+{s('G', LAB)}")
        wn.cell(nr, 14, f"=IF(E{nr}=0,0,1-L{nr}/E{nr})")
        wn.cell(nr, 15, f"=IF(F{nr}=0,0,1-M{nr}/F{nr})")
        wn.cell(nr, 16, f"=(O{nr}-N{nr})*100")
        for c in range(2, 17):
            cell = wn.cell(nr, c)
            cell.font = BOLD if bold else BLACK
            cell.number_format = (PCT if c in (4, 8, 11, 14, 15) else PP if c == 16 else
                                  NUM if c in (5, 6, 7, 12, 13) else "#,##0")
            if bold:
                cell.border = TOP
        if bold:
            wn.cell(nr, 1).border = TOP
        nr += 1
    for L in ("D", "H", "K"):
        wn.conditional_formatting.add(f"{L}6:{L}{nr - 1}", CellIsRule(operator="lessThan", formula=["0"], font=RED))
        wn.conditional_formatting.add(f"{L}6:{L}{nr - 1}", CellIsRule(operator="greaterThan", formula=["0"], font=GREEN))
    wn.cell(nr + 1, 1, "Средний чек = выручка / количество услуг; на него влияют и цены, и структура услуг "
                       "(разложение на объём, цену и структуру будет в факторном анализе).").font = Font(
        name=F, size=9, italic=True)

    # ======================= Тепловая карта =======================
    wh = wb.create_sheet("Тепловая карта", 3)
    wh["A1"], wh["A1"].font = "Тепловая карта отклонений: направление × месяц", TITLE
    wh["A2"], wh["A2"].font = "Клиника (фильтр):", BOLD
    wh["B2"], wh["B2"].fill, wh["B2"].font = "Все", INPUT_FILL, BLUE
    wh["C2"] = '=IF($B$2="Все","*",$B$2)'
    wh["C2"].font = Font(name=F, size=8, color="808080")
    dv3 = DataValidation(type="list", formula1='"Все,C01,C02,C03"', allow_blank=False)
    wh.add_data_validation(dv3)
    dv3.add("B2")
    wh.column_dimensions["A"].width = 30
    for c in range(2, 15):
        wh.column_dimensions[col(c)].width = 10
    scale = lambda: ColorScaleRule(start_type="num", start_value=-0.15, start_color="F8696B",     # noqa: E731
                                   mid_type="num", mid_value=0, mid_color="FFFFFF",
                                   end_type="num", end_value=0.05, end_color="63BE7B")

    def heat(top, title, kind):
        wh.cell(top, 1, title).font = BOLD
        wh.cell(top + 1, 1, "Направление")
        for i, p in enumerate(periods):
            wh.cell(top + 1, 2 + i, p[-2:] + f".{YEAR}")
        wh.cell(top + 1, 14, "Год")
        head(wh, top + 1, 1, 14)
        rr = top + 2
        for code, name in list(DIRS.items()) + [("*", "Все направления")]:
            wh.cell(rr, 1, name).font = BOLD if code == "*" else BLACK
            crit = '"*"' if code == "*" else f'"{code}"'
            for i in range(13):
                if i < 12:
                    per = f'"{periods[i]}"'
                    if kind == "rev":
                        pl = f'SUMIFS({D("F")},{D("A")},{per},{D("E")},"Выручка",{D("D")},{crit},{D("C")},$C$2)'
                        fa = f'SUMIFS({D("G")},{D("A")},{per},{D("E")},"Выручка",{D("D")},{crit},{D("C")},$C$2)'
                    else:
                        pl = f'SUMIFS({Q("E")},{Q("A")},{per},{Q("D")},{crit},{Q("C")},$C$2)'
                        fa = f'SUMIFS({Q("F")},{Q("A")},{per},{Q("D")},{crit},{Q("C")},$C$2)'
                else:
                    if kind == "rev":
                        pl = f'SUMIFS({D("F")},{D("E")},"Выручка",{D("D")},{crit},{D("C")},$C$2)'
                        fa = f'SUMIFS({D("G")},{D("E")},"Выручка",{D("D")},{crit},{D("C")},$C$2)'
                    else:
                        pl = f'SUMIFS({Q("E")},{Q("D")},{crit},{Q("C")},$C$2)'
                        fa = f'SUMIFS({Q("F")},{Q("D")},{crit},{Q("C")},$C$2)'
                cell = wh.cell(rr, 2 + i, f"=IF({pl}=0,\"\",{fa}/{pl}-1)")
                cell.number_format = "0.0%"
                cell.font = BOLD if code == "*" else BLACK
                cell.alignment = Alignment(horizontal="center")
            rr += 1
        wh.conditional_formatting.add(f"B{top + 2}:N{rr - 1}", scale())
        return rr

    nxt = heat(4, "Выручка: факт к плану, %", "rev")
    nxt = heat(nxt + 1, "Объём услуг (штук): факт к плану, %", "qty")
    wh.cell(nxt + 1, 1, "Красное = факт ниже плана. Для клиники C03 и направления «Лазерная коррекция» "
                        "видно провал июль-сентябрь; для стоматологии заметен провал в мае.").font = Font(
        name=F, size=9, italic=True)

    # ======================= Топ отклонений =======================
    wt = wb.create_sheet("Топ отклонений", 4)
    wt["A1"], wt["A1"].font = "Топ отклонений от плана (тыс. ₽)", TITLE
    wt["A2"] = ("Значения рассчитаны скриптом (не формулы) по всей сети; живые расчёты с фильтром по клинике "
                "на листах «План-факт» и «По направлениям».")
    wt["A2"].font = Font(name=F, size=9, italic=True)
    rv = data[data.article == "Выручка"].groupby(["clinic_id", "direction_code", "period"], as_index=False)[
        ["plan", "fact"]].sum()
    qq = qty.rename(columns={"plan_qty": "pq", "fact_qty": "fq"})
    rv = rv.merge(qq[["clinic_id", "direction_code", "period", "pq", "fq"]],
                  on=["clinic_id", "direction_code", "period"], how="left")
    rv["delta"] = rv.fact - rv.plan
    rv["pct"] = rv.fact / rv.plan - 1
    rv["qty_pct"] = rv.fq / rv.pq - 1

    def cell_table(top_row, title, frame):
        wt.cell(top_row, 1, title).font = BOLD
        for i, h in enumerate(["Клиника", "Направление", "Период", "План", "Факт", "Откл.", "Откл., %",
                               "Объём, откл. %"], 1):
            wt.cell(top_row + 1, i, h)
        head(wt, top_row + 1, 1, 8)
        rr = top_row + 2
        for x in frame.itertuples():
            vals = [x.clinic_id, DIRS[x.direction_code], x.period, x.plan, x.fact, x.delta, x.pct, x.qty_pct]
            for i, v in enumerate(vals, 1):
                c = wt.cell(rr, i, v)
                c.font = BLACK
                c.number_format = NUM if i in (4, 5, 6) else "0.0%" if i in (7, 8) else "General"
            rr += 1
        return rr

    nxt = cell_table(4, "1. Самые крупные недоборы выручки: клиника × направление × месяц (10 по сумме)",
                     rv.sort_values("delta").head(10))
    nxt = cell_table(nxt + 1, "2. Самые резкие отклонения в %: ячейки с планом от 2 млн ₽ (10 по проценту)",
                     rv[rv.plan >= 2_000_000].sort_values("pct").head(10))
    wt.cell(nxt, 1, "Как читать: если отклонение объёма почти равно отклонению выручки, причина в количестве услуг "
                    "(событие: праздники, простой оборудования). Если выручка хуже объёма заметно сильнее, "
                    "причина в цене или структуре услуг (косметология: план закладывал индексацию прайса выше факта).").font = Font(
        name=F, size=9, italic=True)
    wt.cell(nxt, 1).alignment = Alignment(wrap_text=True, vertical="top")
    wt.merge_cells(start_row=nxt, start_column=1, end_row=nxt, end_column=8)
    wt.row_dimensions[nxt].height = 40

    nxt += 2
    wt.cell(nxt, 1, "3. Статьи, ухудшающие прибыль, по клиникам за год (только отрицательное влияние)").font = BOLD
    for i, h in enumerate(["Клиника", "Статья", "", "План", "Факт", "Влияние на прибыль", "Откл., %"], 1):
        wt.cell(nxt + 1, i, h)
    head(wt, nxt + 1, 1, 7)
    rr = nxt + 2
    ar = data.groupby(["clinic_id", "article"], as_index=False)[["plan", "fact"]].sum()
    ar["infl"] = ar.apply(lambda x: (x.fact - x.plan) if x.article == "Выручка" else (x.plan - x.fact), axis=1)
    ar["pct"] = ar.fact / ar.plan - 1
    for x in ar[ar.infl < 0].sort_values("infl").itertuples():
        vals = [x.clinic_id, x.article, "", x.plan, x.fact, x.infl, x.pct]
        for i, v in enumerate(vals, 1):
            c = wt.cell(rr, i, v)
            c.font = BLACK
            c.number_format = NUM if i in (4, 5, 6) else "0.0%" if i == 7 else "General"
        rr += 1
    wt.cell(rr + 1, 1, "Расходы, связанные с объёмом (врачи, взносы, материалы), в сумме ниже плана только потому, что "
                       "недобрана выручка; настоящую эффективность затрат покажет факторный анализ на гибком плане.").font = Font(
        name=F, size=9, italic=True)
    for i, w in enumerate([10, 40, 12, 14, 14, 18, 12, 14], 1):
        wt.column_dimensions[col(i)].width = w

    # ======================= Описание =======================
    wi = wb.create_sheet("Описание", 0)
    wi.column_dimensions["A"].width = 120
    text = [
        ("План-факт по P&L сети клиник, 2025", TITLE),
        ("Источники", BOLD),
        ("Факт: проводки 1С после исправлений сверки (тот же расчёт, что в PnL_2025.xlsx, итоги совпадают). "
         "План: budget_revenue.csv (количество, плановая цена, плановая себестоимость единицы по услугам) и "
         "budget_opex.csv (расходы по клиникам и статьям). Объём услуг: visits.csv и visit_items.csv.", BLACK),
        ("Как читать", BOLD),
        ("Отклонение = факт − план. Влияние на прибыль: плюс выгодно (выручка выше плана, расход ниже плана), "
         "минус невыгодно. Фильтр клиники на листах «План-факт», «По направлениям» и «Тепловая карта» "
         "(жёлтая ячейка B2).", BLACK),
        ("Правила построения плана", BOLD),
        ("Плановая выручка = плановое количество × плановая цена по услугам. Плановые материалы и лаборатории = "
         "плановое количество × плановая себестоимость единицы. Остальные статьи заданы бюджетом по клиникам "
         "(оплата труда врачей считается от плановой выручки по средней доле врачей направления).", BLACK),
        ("Ограничения", BOLD),
        ("Косвенные расходы в плане не разнесены по направлениям, поэтому на листе «По направлениям» "
         "показана прямая маржа (после материалов и лабораторий), а полный P&L только по клиникам.", BLACK),
        ("Отклонение по оплате труда врачей включает эффект структуры: план считает среднюю долю по направлению, "
         "факт зависит от грейдов реально принимавших врачей.", BLACK),
        ("Налог на прибыль и НДС не моделируются: P&L заканчивается прибылью до налога.", BLACK),
        ("Следующий шаг: факторный анализ маржи (объём, цена, структура, себестоимость) разложит отклонения "
         "на причины; этот файл показывает, где и на сколько.", BLACK),
    ]
    for i, (t, f) in enumerate(text, 1):
        c = wi.cell(i, 1, t)
        c.font, c.alignment = f, Alignment(wrap_text=True, vertical="top")
    return wb


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    data, qty = prepare()
    wb = build(data, qty)
    path = OUT / "PlanFact_2025.xlsx"
    wb.save(path)
    recalculate(path)
    print(f"Сохранено: {path}  (строк данных: {len(data)}, объёмов: {len(qty)})")

    t = data.groupby("article")[["plan", "fact"]].sum()
    rev_p, rev_f = t.loc["Выручка"]
    cost_a = [a for a in t.index if a != "Выручка"]
    prof_p = rev_p - t.loc[cost_a, "plan"].sum()
    prof_f = rev_f - t.loc[cost_a, "fact"].sum()
    print(f"Контроль в pandas: выручка план {rev_p:,.0f} факт {rev_f:,.0f} ({rev_f / rev_p - 1:.1%}); "
          f"прибыль до налога план {prof_p:,.0f} факт {prof_f:,.0f} (откл. {prof_f - prof_p:,.0f})")


if __name__ == "__main__":
    main()

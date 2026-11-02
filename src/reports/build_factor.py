"""
Шаг 11: факторный анализ прямой маржи (план -> факт) за 2025 год в Excel.

Прямая маржа = выручка - материалы - лаборатории и внешние услуги: ровно те статьи,
которые в плане заданы на уровне услуг (количество, цена, себестоимость единицы).
Шаг 2 (достройка до валовой маржи: врачи, медперсонал, взносы, амортизация) описан ниже и
вынесен на листы «Мост валовая», «Срезы валовая», «Данные ФОТ».

Вход:  data/corrected/ - visits, visit_items, services (факт по услугам),
                         budget_revenue (план по услугам), materials_1c (факт затрат из 1С)
Выход: data/reports/FactorAnalysis_2025.xlsx

Метод. Для каждой строки «месяц x клиника x услуга» маржа единицы m = цена - материалы - лаборатории.
Отклонение факта от плана раскладывается точно, без остатка, на 6 факторов:

  Объём             = плановая маржа ячейки x (индекс объёма - 1); индекс объёма считается
                      в плановых ценах по ячейке «месяц x клиника x направление»:
                      V = сумма(факт. кол-во x плановая цена) / сумма(плановое кол-во x плановая цена)
  Структура услуг   = сумма(факт. кол-во x плановая маржа единицы) - V x плановая маржа ячейки
                      (сдвиг состава услуг внутри направления в сторону более или менее маржинальных)
  Цена: прайс       = сумма(факт. кол-во x действовавший прайс) - сумма(факт. кол-во x плановая цена)
                      (план закладывал одну индексацию, факт - другую)
  Цена: скидки      = факт. выручка - сумма(факт. кол-во x действовавший прайс)
  Материалы         = -(факт. материалы - факт. кол-во x плановые материалы единицы)
  Лаборатории       = -(факт. лаборатории - факт. кол-во x плановые лаборатории единицы)

Плюс = выгодно для маржи, минус = невыгодно. Книга живая: «Данные» содержат входные числа и
формулы эффектов по строкам, отчётные листы считаются SUMIFS.

Факт затрат: 1С хранит материалы и лаборатории по направлениям. На услуги они разнесены
пропорционально нормативу (количество x норма единицы), поэтому итоги совпадают с 1С.

Достройка до валовой маржи. Оплата врачей считается на гибком плане (плановая доля врачей
направления x фактическая выручка) и раскладывается на три эффекта, взносы внутри строк ФОТ:
  Врачи: объём       = -(факт.кол-во x план.цена - план.выручка) x доля x (1 + ставка взносов)
  Врачи: цена/скидки = -(факт.выручка - факт.кол-во x план.цена) x доля x (1 + ставка взносов)
  Врачи: грейды/доля = факт.выручка x доля x (1 + ставка) - (факт. оплата + факт. взносы)
Медперсонал и амортизация заданы в плане по клиникам, поэтому идут отдельной строкой «—».
Ставка взносов в плане 22 процента - допущение демо (плоская, без предельной базы).

Запуск:  python src/reports/build_factor.py
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.recalc import recalculate  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "reports"
YEAR = 2025
CLINICS = ["C01", "C02", "C03"]
DIRS = {"DENT": "Стоматология", "FAM": "Семейная медицина", "COSM": "Косметология",
        "PLAS": "Пластическая хирургия", "LASER": "Лазерная коррекция зрения"}
CELL = ["period", "clinic_id", "direction_code"]

FACTORS = [("Объём", "W"), ("Структура услуг", "X"), ("Цена: прайс (индексация)", "Y"),
           ("Цена: скидки", "Z"), ("Материалы", "AA"), ("Лаборатории и внешние услуги", "AB")]

F = "Arial"
BLACK, BOLD = Font(name=F, size=10), Font(name=F, size=10, bold=True)
BLUE, GREEN_LINK = Font(name=F, size=10, color="0000FF"), Font(name=F, size=10, color="008000")
RED, GREEN = Font(name=F, size=10, color="C00000"), Font(name=F, size=10, color="00803C")
TITLE, NOTE = Font(name=F, size=14, bold=True), Font(name=F, size=9, italic=True, color="595959")
HDR_FONT = Font(name=F, size=10, bold=True, color="FFFFFF")
HDR_FILL, SEC_FILL = PatternFill("solid", fgColor="1F3864"), PatternFill("solid", fgColor="D9E2F3")
INPUT_FILL = PatternFill("solid", fgColor="FFFF99")
TOP = Border(top=Side(style="thin"))
NUM = "#,##0,;(#,##0,);-"
PCT = "0.0%;(0.0%);-"
PP = '0.0" п.п.";(0.0" п.п.");-'


# ------------------------------------------------------------------ данные
def prepare():
    visits = pd.read_csv(CORR / "visits.csv")
    items = pd.read_csv(CORR / "visit_items.csv")
    services = pd.read_csv(CORR / "services.csv")
    plan = pd.read_csv(CORR / "budget_revenue.csv")
    mat1c = pd.read_csv(CORR / "materials_1c.csv")

    im = items.merge(visits[["visit_id", "visit_date", "clinic_id"]], on="visit_id").merge(
        services[["service_id", "service_name", "material_cost", "external_cost"]], on="service_id")
    im["period"] = im.visit_date.str[:7]
    im["list_total"] = im.list_price_at_date * im.qty
    im["norm_mat"] = im.qty * im.material_cost
    im["norm_ext"] = im.qty * im.external_cost
    fact = im.groupby(["period", "clinic_id", "service_id"], as_index=False).agg(
        fact_qty=("qty", "sum"), fact_revenue=("revenue", "sum"), list_total=("list_total", "sum"),
        norm_mat=("norm_mat", "sum"), norm_ext=("norm_ext", "sum"))

    d = plan.merge(fact, on=["period", "clinic_id", "service_id"], how="outer", validate="1:1")
    if d[["plan_qty", "fact_qty"]].isna().any().any():
        raise ValueError("Ячейки плана и факта не совпадают: нужен разбор")
    d = d.merge(services[["service_id", "service_name"]], on="service_id")
    d["direction_code"] = d.direction_code.fillna("")

    # факт затрат 1С (по направлениям) -> по услугам пропорционально нормативу
    cell_norm = d.groupby(CELL)[["norm_mat", "norm_ext"]].transform("sum")
    m = mat1c[CELL + ["material_fact", "external_fact"]]
    d = d.merge(m, on=CELL, how="left")
    d["mat_fact"] = np.where(cell_norm.norm_mat > 0, d.material_fact * d.norm_mat / cell_norm.norm_mat.replace(0, 1), 0.0)
    d["ext_fact"] = np.where(cell_norm.norm_ext > 0, d.external_fact * d.norm_ext / cell_norm.norm_ext.replace(0, 1), 0.0)

    d = d.sort_values(["period", "clinic_id", "direction_code", "service_id"]).reset_index(drop=True)
    d = d[["period", "clinic_id", "direction_code", "service_id", "service_name", "plan_qty", "fact_qty",
           "plan_price", "list_total", "fact_revenue", "plan_unit_material", "plan_unit_external",
           "mat_fact", "ext_fact"]]
    for c in ("list_total", "fact_revenue", "mat_fact", "ext_fact"):
        d[c] = d[c].round(2)
    return d, mat1c


def gross_profit_1c():
    """Валовая прибыль по проводкам 1С: выручка 90.01 минус дебетовые обороты счёта 20 (без закрытия месяца)."""
    post = pd.read_csv(CORR / "postings_1c.csv", dtype={"debit": str, "credit": str})
    post = post[post.doc_type != "Закрытие месяца"]
    return float(post[post.credit == "90.01"].amount.sum() - post[post.debit == "20"].amount.sum())


def effects(d):
    """То же разложение в pandas: контроль формул Excel и таблицы «Топ факторов»."""
    e = d.copy()
    e["q_f_pp"] = e.fact_qty * e.plan_price
    e["q_p_pp"] = e.plan_qty * e.plan_price
    e["m_unit"] = e.plan_price - e.plan_unit_material - e.plan_unit_external
    e["plan_margin"] = e.plan_qty * e.m_unit
    e["vol_idx"] = e.groupby(CELL).q_f_pp.transform("sum") / e.groupby(CELL).q_p_pp.transform("sum")
    e["volume"] = e.plan_margin * (e.vol_idx - 1)
    e["mix"] = e.fact_qty * e.m_unit - e.vol_idx * e.plan_margin
    e["price_list"] = e.list_total - e.q_f_pp
    e["discount"] = e.fact_revenue - e.list_total
    e["materials"] = -(e.mat_fact - e.fact_qty * e.plan_unit_material)
    e["labs"] = -(e.ext_fact - e.fact_qty * e.plan_unit_external)
    e["fact_margin"] = e.fact_revenue - e.mat_fact - e.ext_fact
    e["check"] = e.fact_margin - (e.plan_margin + e.volume + e["mix"] + e.price_list + e.discount + e.materials + e.labs)
    e["plan_revenue"] = e.q_p_pp
    return e


# ------------------------------------------------------------------ оформление
def head(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def sign_format(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=RED))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=GREEN))


def build(d):
    e = effects(d)
    wb = Workbook()
    n = len(d)
    nd = n + 1

    # ---------------------------------------------------------------- Данные
    wd = wb.active
    wd.title = "Данные"
    hdr = ["Период", "Клиника", "Направление", "Услуга", "Название", "Кол-во, план", "Кол-во, факт",
           "Цена, план", "Выручка факт по прайсу", "Выручка, факт", "Материалы ед., план", "Лаборатории ед., план",
           "Материалы, факт", "Лаборатории, факт",
           "Ключ ячейки", "Факт кол-во × план цена", "План выручка", "Маржа ед., план", "Маржа, план",
           "Факт кол-во × маржа ед. план", "Индекс объёма ячейки",
           "Маржа, план (ф)", "Объём", "Структура услуг", "Цена: прайс", "Цена: скидки", "Материалы", "Лаборатории",
           "Маржа, факт", "Проверка"]
    # буквы: A..N вход, O ключ, P..T помощники, U индекс, V план маржа, W..AB факторы, AC факт, AD проверка
    for i, h in enumerate(hdr, 1):
        wd.cell(1, i, h)
    head(wd, 1, 1, len(hdr))
    wd.row_dimensions[1].height = 42
    inp = d.values.tolist()
    for r, row in enumerate(inp, 2):
        for c, v in enumerate(row, 1):
            cell = wd.cell(r, c, v)
            cell.font = BLUE if c >= 6 else BLACK
        wd.cell(r, 15, f'=A{r}&"|"&B{r}&"|"&C{r}')
        wd.cell(r, 16, f"=G{r}*H{r}")
        wd.cell(r, 17, f"=F{r}*H{r}")
        wd.cell(r, 18, f"=H{r}-K{r}-L{r}")
        wd.cell(r, 19, f"=F{r}*R{r}")
        wd.cell(r, 20, f"=G{r}*R{r}")
        wd.cell(r, 21, f"=SUMIFS($P$2:$P${nd},$O$2:$O${nd},O{r})/SUMIFS($Q$2:$Q${nd},$O$2:$O${nd},O{r})")
        wd.cell(r, 22, f"=S{r}")
        wd.cell(r, 23, f"=S{r}*(U{r}-1)")
        wd.cell(r, 24, f"=T{r}-U{r}*S{r}")
        wd.cell(r, 25, f"=I{r}-P{r}")
        wd.cell(r, 26, f"=J{r}-I{r}")
        wd.cell(r, 27, f"=-(M{r}-G{r}*K{r})")
        wd.cell(r, 28, f"=-(N{r}-G{r}*L{r})")
        wd.cell(r, 29, f"=J{r}-M{r}-N{r}")
        wd.cell(r, 30, f"=AC{r}-(V{r}+W{r}+X{r}+Y{r}+Z{r}+AA{r}+AB{r})")
        for c in range(15, 31):
            wd.cell(r, c).font = BLACK
        for c in list(range(6, 15)) + list(range(16, 21)) + list(range(22, 31)):
            wd.cell(r, c).number_format = "#,##0.00"
        wd.cell(r, 21).number_format = "0.0000"
    widths = [9, 8, 11, 8, 36, 10, 10, 10, 14, 14, 11, 11, 13, 13, 22, 14, 14, 11, 14, 14, 11,
              14, 13, 13, 13, 13, 13, 13, 14, 10]
    for i, w in enumerate(widths, 1):
        wd.column_dimensions[col(i)].width = w
    wd.freeze_panes = "F2"
    wd.auto_filter.ref = f"A1:AD{nd}"
    D = lambda c: f"'Данные'!${c}$2:${c}${nd}"  # noqa: E731

    periods = [f"{YEAR}-{m:02d}" for m in range(1, 13)]

    def crit(clinic="$C$2", direction="$C$3"):
        parts = []
        if clinic:
            parts.append(f"{D('B')},{clinic}")
        if direction:
            parts.append(f"{D('C')},{direction}")
        return ",".join(parts)

    # ---------------------------------------------------------------- Мост
    ws = wb.create_sheet("Мост", 0)
    ws["A1"], ws["A1"].font = "Факторный анализ прямой маржи, 2025 (тыс. ₽)", TITLE
    ws["A2"], ws["A2"].font = "Клиника (фильтр):", BOLD
    ws["A3"], ws["A3"].font = "Направление (фильтр):", BOLD
    for ref, wild, vals in (("B2", "C2", CLINICS), ("B3", "C3", list(DIRS))):
        ws[ref], ws[ref].fill, ws[ref].font = "Все", INPUT_FILL, BLUE
        ws[wild] = f'=IF($B${ref[1]}="Все","*",$B${ref[1]})'
        ws[wild].font = Font(name=F, size=8, color="808080")
        dv = DataValidation(type="list", formula1='"' + ",".join(["Все"] + vals) + '"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ref)
    ws["D2"], ws["D3"] = "← Все, C01, C02 или C03", "← Все, DENT, FAM, COSM, PLAS или LASER"
    ws["D2"].font = ws["D3"].font = NOTE
    ws["A4"] = "Плюс = выгодно для маржи, минус = невыгодно. Прямая маржа = выручка − материалы − лаборатории."
    ws["A4"].font = NOTE
    for i, h in enumerate(["Показатель", "Сумма", "% плановой маржи", "Что означает"], 1):
        ws.cell(6, i, h)
    head(ws, 6, 1, 8)
    ws["D6"].alignment = Alignment(horizontal="left", vertical="center")
    S = lambda c: f"=SUMIFS({D(c)},{crit()})"  # noqa: E731
    meaning = {
        "Объём": "Больше или меньше услуг, чем в плане, при плановых ценах и составе",
        "Структура услуг": "Состав услуг внутри направления сместился к более или менее маржинальным",
        "Цена: прайс (индексация)": "Действовавший прайс отличается от плановой цены (другая индексация)",
        "Цена: скидки": "Скидки пациентам снизили факт. цену относительно прайса",
        "Материалы": "Материалы на единицу услуги дороже или дешевле плана",
        "Лаборатории и внешние услуги": "Лаборатории и внешние услуги на единицу дороже или дешевле плана"}
    ws.cell(7, 1, "Плановая прямая маржа").font = BOLD
    ws.cell(7, 2, S("V")).font = BOLD
    r = 8
    for name, c in FACTORS:
        ws.cell(r, 1, name).font = BLACK
        ws.cell(r, 2, S(c)).font = GREEN_LINK
        ws.cell(r, 3, f"=IF($B$7=0,0,B{r}/$B$7)")
        ws.cell(r, 4, meaning[name]).font = NOTE
        r += 1
    fr = r  # строка факта
    ws.cell(fr, 1, "Фактическая прямая маржа").font = BOLD
    ws.cell(fr, 2, f"=B7+SUM(B8:B{fr - 1})").font = BOLD
    ws.cell(fr, 3, f"=IF($B$7=0,0,B{fr}/$B$7-1)")
    ws.cell(fr + 1, 1, "Контроль: факт маржи по данным минус факт по мосту").font = NOTE
    ws.cell(fr + 1, 2, f"=SUMIFS({D('AC')},{crit()})-B{fr}").font = NOTE
    ws.cell(fr + 1, 4, "должно быть 0").font = NOTE
    for rr in range(7, fr + 2):
        ws.cell(rr, 2).number_format = NUM if rr != fr + 1 else "#,##0.00"
        ws.cell(rr, 3).number_format = PCT
    for c in range(1, 5):
        ws.cell(fr, c).border = TOP
        ws.cell(7, c).fill = SEC_FILL
    sign_format(ws, f"B8:C{fr - 1}")
    ws.cell(fr, 3).font = BOLD

    q0 = fr + 3
    for i, h in enumerate(["", "План", "Факт", "Откл."], 1):
        ws.cell(q0, i, h)
    head(ws, q0, 1, 4)
    ws.cell(q0 + 1, 1, "Выручка")
    ws.cell(q0 + 1, 2, f"=SUMIFS({D('Q')},{crit()})")
    ws.cell(q0 + 1, 3, f"=SUMIFS({D('J')},{crit()})")
    ws.cell(q0 + 2, 1, "Прямая маржа")
    ws.cell(q0 + 2, 2, "=B7")
    ws.cell(q0 + 2, 3, f"=B{fr}")
    ws.cell(q0 + 3, 1, "Прямая маржа, % выручки")
    for c in (2, 3):
        L = col(c)
        ws.cell(q0 + 3, c, f"=IF({L}{q0 + 1}=0,0,{L}{q0 + 2}/{L}{q0 + 1})")
    ws.cell(q0 + 1, 4, f"=C{q0 + 1}-B{q0 + 1}")
    ws.cell(q0 + 2, 4, f"=C{q0 + 2}-B{q0 + 2}")
    ws.cell(q0 + 3, 4, f"=(C{q0 + 3}-B{q0 + 3})*100")
    for rr in (q0 + 1, q0 + 2):
        for c in (2, 3, 4):
            ws.cell(rr, c).number_format = NUM
    for c in (2, 3):
        ws.cell(q0 + 3, c).number_format = PCT
    ws.cell(q0 + 3, 4).number_format = PP
    for rr in range(q0 + 1, q0 + 4):
        for c in range(1, 5):
            ws.cell(rr, c).font = BLACK
    sign_format(ws, f"D{q0 + 1}:D{q0 + 3}")
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    for c_ in "EFGH":
        ws.column_dimensions[c_].width = 14

    ch = BarChart()
    ch.type, ch.title = "bar", "Влияние факторов на прямую маржу, тыс. ₽"
    ch.add_data(Reference(ws, min_col=2, min_row=8, max_row=fr - 1), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=1, min_row=8, max_row=fr - 1))
    ch.legend = None
    ch.series[0].invertIfNegative = False
    ch.series[0].graphicalProperties.solidFill = "2E75B6"
    ch.y_axis.numFmt = "#,##0,"
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    ch.x_axis.scaling.orientation = "maxMin"
    ch.x_axis.tickLblPos = "low"
    ch.height, ch.width = 8.5, 22
    ws.add_chart(ch, f"A{q0 + 6}")

    # ---------------------------------------------------------------- По месяцам
    wm = wb.create_sheet("По месяцам", 1)
    wm["A1"], wm["A1"].font = "Факторы по месяцам (тыс. ₽)", TITLE
    wm["A2"], wm["A2"].font = "Клиника (фильтр):", BOLD
    wm["A3"], wm["A3"].font = "Направление (фильтр):", BOLD
    for ref, wild, vals in (("B2", "C2", CLINICS), ("B3", "C3", list(DIRS))):
        wm[ref], wm[ref].fill, wm[ref].font = "Все", INPUT_FILL, BLUE
        wm[wild] = f'=IF($B${ref[1]}="Все","*",$B${ref[1]})'
        wm[wild].font = Font(name=F, size=8, color="808080")
        dv = DataValidation(type="list", formula1='"' + ",".join(["Все"] + vals) + '"', allow_blank=False)
        wm.add_data_validation(dv)
        dv.add(ref)
    wm["A4"] = "Объём в мае по стоматологии и в июле-сентябре по лазеру клиники C03: выбери направление и клинику."
    wm["A4"].font = NOTE
    wm.cell(6, 1, "Фактор")
    for i, p in enumerate(periods):
        wm.cell(6, 2 + i, p)
    wm.cell(6, 14, f"Итого {YEAR}")
    head(wm, 6, 1, 14)
    rr = 7
    wm.cell(rr, 1, "Плановая прямая маржа").font = BOLD
    for i in range(12):
        L = col(2 + i)
        wm.cell(rr, 2 + i, f"=SUMIFS({D('V')},{D('A')},{L}$6,{crit()})").font = BOLD
    wm.cell(rr, 14, "=SUM(B7:M7)").font = BOLD
    rr += 1
    for name, c in FACTORS:
        wm.cell(rr, 1, name).font = BLACK
        for i in range(12):
            L = col(2 + i)
            wm.cell(rr, 2 + i, f"=SUMIFS({D(c)},{D('A')},{L}$6,{crit()})")
        wm.cell(rr, 14, f"=SUM(B{rr}:M{rr})")
        rr += 1
    wm.cell(rr, 1, "Фактическая прямая маржа").font = BOLD
    for c in range(2, 15):
        L = col(c)
        wm.cell(rr, c, f"={L}7+SUM({L}8:{L}{rr - 1})").font = BOLD
        wm.cell(rr, c).border = TOP
    wm.cell(rr, 1).border = TOP
    for r2 in range(7, rr + 1):
        for c in range(2, 15):
            wm.cell(r2, c).number_format = NUM
            if wm.cell(r2, c).font.bold is not True:
                wm.cell(r2, c).font = BLACK
    sign_format(wm, f"B8:N{rr - 1}")
    wm.cell(rr + 1, 1, "Объём к плану, % (по выручке в плановых ценах)").font = BOLD
    for i in range(12):
        L = col(2 + i)
        wm.cell(rr + 1, 2 + i, f"=IF(SUMIFS({D('Q')},{D('A')},{L}$6,{crit()})=0,0,"
                               f"SUMIFS({D('P')},{D('A')},{L}$6,{crit()})/SUMIFS({D('Q')},{D('A')},{L}$6,{crit()})-1)")
        wm.cell(rr + 1, 2 + i).number_format = PCT
    wm.cell(rr + 1, 14, f"=IF(SUMIFS({D('Q')},{crit()})=0,0,SUMIFS({D('P')},{crit()})/SUMIFS({D('Q')},{crit()})-1)")
    wm.cell(rr + 1, 14).number_format = PCT
    from openpyxl.formatting.rule import ColorScaleRule
    wm.conditional_formatting.add(f"B{rr + 1}:N{rr + 1}", ColorScaleRule(
        start_type="num", start_value=-0.15, start_color="F8696B", mid_type="num", mid_value=0,
        mid_color="FFFFFF", end_type="num", end_value=0.05, end_color="63BE7B"))
    wm.column_dimensions["A"].width = 44
    for c in range(2, 15):
        wm.column_dimensions[col(c)].width = 11.5
    wm.freeze_panes = "B7"

    # ---------------------------------------------------------------- Срезы
    wc = wb.create_sheet("Срезы", 2)
    wc["A1"], wc["A1"].font = "Факторы по направлениям и клиникам, 2025 (тыс. ₽)", TITLE
    cols_h = ["План маржи", "Объём", "Структура услуг", "Цена: прайс", "Цена: скидки", "Материалы",
              "Лаборатории", "Факт маржи", "Откл. маржи", "Маржа план, %", "Маржа факт, %", "Откл., п.п."]
    srcs = ["V", "W", "X", "Y", "Z", "AA", "AB", "AC"]

    def slice_table(top, title, key_col, keys, names, extra_crit, filter_label, filter_ref, filter_wild, vals, filter_col):
        wc.cell(top, 1, title).font = BOLD
        wc.cell(top + 1, 1, filter_label).font = BOLD
        wc[filter_ref], wc[filter_ref].fill, wc[filter_ref].font = "Все", INPUT_FILL, BLUE
        wc[filter_wild] = f'=IF(${filter_ref[0]}${filter_ref[1:]}="Все","*",${filter_ref[0]}${filter_ref[1:]})'
        wc[filter_wild].font = Font(name=F, size=8, color="808080")
        dv = DataValidation(type="list", formula1='"' + ",".join(["Все"] + vals) + '"', allow_blank=False)
        wc.add_data_validation(dv)
        dv.add(filter_ref)
        h = top + 3
        wc.cell(h, 1, "Срез")
        for i, t in enumerate(cols_h):
            wc.cell(h, 2 + i, t)
        head(wc, h, 1, 13)
        wc.row_dimensions[h].height = 30
        r0 = h + 1
        for k, nm in zip(keys + ["*"], names + ["Итого"]):
            rr = h + 1 + (keys + ["*"]).index(k)
            bold = k == "*"
            wc.cell(rr, 1, nm).font = BOLD if bold else BLACK
            wc.cell(rr, 15, k).font = Font(name=F, size=8, color="808080")
            for i, s in enumerate(srcs):
                wc.cell(rr, 2 + i, f"=SUMIFS({D(s)},{D(key_col)},$O{rr},{D(filter_col)},${filter_wild[0]}${filter_wild[1:]})")
                wc.cell(rr, 2 + i).number_format = NUM
                wc.cell(rr, 2 + i).font = BOLD if bold else BLACK
            wc.cell(rr, 10, f"=I{rr}-B{rr}")
            wc.cell(rr, 11, f"=IF(SUMIFS({D('Q')},{D(key_col)},$O{rr},{D(filter_col)},${filter_wild[0]}${filter_wild[1:]})=0,0,B{rr}/SUMIFS({D('Q')},{D(key_col)},$O{rr},{D(filter_col)},${filter_wild[0]}${filter_wild[1:]}))")
            wc.cell(rr, 12, f"=IF(SUMIFS({D('J')},{D(key_col)},$O{rr},{D(filter_col)},${filter_wild[0]}${filter_wild[1:]})=0,0,I{rr}/SUMIFS({D('J')},{D(key_col)},$O{rr},{D(filter_col)},${filter_wild[0]}${filter_wild[1:]}))")
            wc.cell(rr, 13, f"=(L{rr}-K{rr})*100")
            wc.cell(rr, 10).number_format = NUM
            wc.cell(rr, 11).number_format = wc.cell(rr, 12).number_format = PCT
            wc.cell(rr, 13).number_format = PP
            for c in (10, 11, 12, 13):
                wc.cell(rr, c).font = BOLD if bold else BLACK
            if bold:
                for c in range(1, 14):
                    wc.cell(rr, c).border = TOP
        last = h + len(keys) + 1
        sign_format(wc, f"C{h + 1}:H{last}")
        sign_format(wc, f"J{h + 1}:J{last}")
        sign_format(wc, f"M{h + 1}:M{last}")
        return last

    last = slice_table(2, "По направлениям", "C", list(DIRS), list(DIRS.values()), None,
                       "Клиника (фильтр):", "B3", "C3", CLINICS, "B")
    slice_table(last + 3, "По клиникам", "B", CLINICS, ["Невский", "Север", "Юг"], None,
                "Направление (фильтр):", f"B{last + 4}", f"C{last + 4}", list(DIRS), "C")
    wc.column_dimensions["A"].width = 30
    for c in range(2, 14):
        wc.column_dimensions[col(c)].width = 13
    wc.column_dimensions["O"].width = 6
    wc.freeze_panes = "B5"

    # ---------------------------------------------------------------- По услугам
    wsv = wb.create_sheet("По услугам", 3)
    wsv["A1"], wsv["A1"].font = "Факторы по услугам, 2025 (тыс. ₽)", TITLE
    wsv["A2"], wsv["A2"].font = "Клиника (фильтр):", BOLD
    wsv["B2"], wsv["B2"].fill, wsv["B2"].font = "Все", INPUT_FILL, BLUE
    wsv["C2"] = '=IF($B$2="Все","*",$B$2)'
    wsv["C2"].font = Font(name=F, size=8, color="808080")
    dv = DataValidation(type="list", formula1='"Все,C01,C02,C03"', allow_blank=False)
    wsv.add_data_validation(dv)
    dv.add("B2")
    svc = d[["service_id", "service_name", "direction_code"]].drop_duplicates().sort_values("service_id")
    for i, h in enumerate(["Услуга", "Название", "Напр."] + cols_h[:8], 1):
        wsv.cell(4, i, h)
    head(wsv, 4, 1, 11)
    wsv.row_dimensions[4].height = 30
    for j, s in enumerate(svc.itertuples(), 5):
        wsv.cell(j, 1, s.service_id).font = BLACK
        wsv.cell(j, 2, s.service_name).font = BLACK
        wsv.cell(j, 3, s.direction_code).font = BLACK
        for i, src in enumerate(srcs):
            wsv.cell(j, 4 + i, f"=SUMIFS({D(src)},{D('D')},$A{j},{D('B')},$C$2)")
            wsv.cell(j, 4 + i).number_format = NUM
            wsv.cell(j, 4 + i).font = BLACK
    sign_format(wsv, f"E5:J{4 + len(svc)}")
    wsv.column_dimensions["A"].width = 9
    wsv.column_dimensions["B"].width = 42
    wsv.column_dimensions["C"].width = 8
    for c in range(4, 12):
        wsv.column_dimensions[col(c)].width = 13
    wsv.freeze_panes = "D5"

    # ---------------------------------------------------------------- Топ факторов (значения)
    wt = wb.create_sheet("Топ факторов", 4)
    wt["A1"], wt["A1"].font = "Где сильнее всего сработали факторы (тыс. ₽)", TITLE
    wt["A2"] = "Значения рассчитаны скриптом по всей сети (не формулы). Фильтры живых листов: «Мост», «По месяцам», «Срезы»."
    wt["A2"].font = NOTE
    cell_e = e.groupby(CELL, as_index=False).agg(
        plan_margin=("plan_margin", "sum"), volume=("volume", "sum"), mix=("mix", "sum"),
        price_list=("price_list", "sum"), discount=("discount", "sum"), materials=("materials", "sum"),
        labs=("labs", "sum"), q_f_pp=("q_f_pp", "sum"), q_p_pp=("q_p_pp", "sum"))
    cell_e["vol_pct"] = cell_e.q_f_pp / cell_e.q_p_pp - 1

    def top_table(top, title, frame, extra=None):
        wt.cell(top, 1, title).font = BOLD
        hh = ["Период", "Клиника", "Направление", "Эффект", "Объём к плану, %"]
        for i, h in enumerate(hh, 1):
            wt.cell(top + 1, i, h)
        head(wt, top + 1, 1, len(hh))
        rr = top + 2
        for x in frame.itertuples():
            vals = [x.period, x.clinic_id, DIRS[x.direction_code], x.eff, x.vol_pct]
            for i, v in enumerate(vals, 1):
                cc = wt.cell(rr, i, v)
                cc.font = BLACK
                cc.number_format = NUM if i == 4 else PCT if i == 5 else "General"
            rr += 1
        return rr + 1

    nxt = top_table(4, "Объём: 10 самых слабых ячеек (месяц × клиника × направление)",
                    cell_e.assign(eff=cell_e.volume).sort_values("eff").head(10))
    nxt = top_table(nxt, "Цена, прайс: 10 самых невыгодных ячеек",
                    cell_e.assign(eff=cell_e.price_list).sort_values("eff").head(10))
    nxt = top_table(nxt, "Скидки: 10 самых больших ячеек",
                    cell_e.assign(eff=cell_e.discount).sort_values("eff").head(10))
    nxt = top_table(nxt, "Материалы: 10 самых невыгодных ячеек",
                    cell_e.assign(eff=cell_e.materials).sort_values("eff").head(10))
    for i, w in enumerate([10, 10, 30, 14, 16], 1):
        wt.column_dimensions[col(i)].width = w

    # ---------------------------------------------------------------- Описание
    wi = wb.create_sheet("Описание", 0)
    wi.column_dimensions["A"].width = 125
    text = [
        ("Факторный анализ маржи, 2025", TITLE),
        ("Что раскладываем", BOLD),
        ("Отклонение факта от плана по прямой марже (выручка − материалы − лаборатории). Именно эти статьи заданы в плане "
         "на уровне услуг: количество, цена и себестоимость единицы. Врачи, медперсонал и взносы в плане услугам не "
         "приписаны: они добавлены отдельно на листах «Мост валовая» и «Срезы валовая» (достройка до валовой маржи).", BLACK),
        ("Факторы (разложение точное, остатка нет; контроль на листе «Мост»)", BOLD),
        ("Объём: плановая маржа ячейки × (индекс объёма − 1). Индекс объёма считается в плановых ценах по ячейке "
         "«месяц × клиника × направление», поэтому дешёвые и дорогие услуги не складываются штуками.", BLACK),
        ("Структура услуг: сдвиг состава услуг внутри направления к более или менее маржинальным.", BLACK),
        ("Цена: прайс: разница между действовавшим прайсом и плановой ценой (план мог заложить другую индексацию).", BLACK),
        ("Цена: скидки: разница между фактической выручкой и выручкой по действовавшему прайсу.", BLACK),
        ("Материалы и лаборатории: фактическая стоимость единицы против плановой, умноженная на фактическое количество.", BLACK),
        ("Достройка до валовой маржи (листы «Мост валовая», «Срезы валовая», «Данные ФОТ»)", BOLD),
        ("Оплата врачей считается на гибком плане: плановая доля врачей направления × фактическая выручка. Три эффекта: "
         "объём и цена/скидки показывают, что оплата врачей снижается вместе с выручкой (это следствие недобора, а не "
         "экономия); «структура грейдов и доля» это отклонение фактической оплаты от гибкого плана, то есть изменение "
         "состава грейдов и долей. Это не вердикт о переплате: более дорогие грейды могут приводить более высокий чек.", BLACK),
        ("Взносы включены в строки ФОТ. Ставка в плане 22% (ячейка Z1 на листе «Данные ФОТ») — ДОПУЩЕНИЕ ДЕМО: плоская "
         "эффективная ставка без учёта предельной базы. В реальной клинике к концу года ставка падает, и появился бы "
         "отдельный эффект лимита.", BLACK),
        ("Медперсонал и амортизация в 1С разнесены по направлениям, а в плане заданы только по клиникам. Поэтому в мосте "
         "они показаны на уровне клиники, а на срезе по направлениям идут строкой «Не распределено». При выбранном "
         "направлении в мосте они не учитываются.", BLACK),
        ("Источники", BOLD),
        ("Факт по услугам: visit_items и visits после исправлений сверки. План: budget_revenue. "
         "Затраты факта: materials_1c (1С хранит их по направлениям); на услуги разнесены пропорционально нормативу "
         "(количество × норма единицы), поэтому суммы совпадают с 1С.", BLACK),
        ("Ограничения", BOLD),
        ("1. Структура направлений и клиник отдельным фактором не выделена: она входит в объём, потому что объём "
         "считается по каждой ячейке «месяц × клиника × направление» и суммируется.", BLACK),
        ("2. Фактор «Материалы/Лаборатории» показывает цену единицы. Расход на единицу (норматив) в данных не меняется, "
         "поэтому перерасход нормы от роста цен отделить нельзя.", BLACK),
        ("3. Эффект объёма смешивает фоновый недобор плана и разовые события. Разовые события видны на листах "
         "«По месяцам» и «Топ факторов».", BLACK),
        ("4. Налог на прибыль и НДС не моделируются.", BLACK),
        ("5. Валовая маржа в мосте сверена с 1С по сети в целом; на срезах по клиникам и направлениям сверка с ОСВ не делается.", BLACK),
        ("Цвета: синий шрифт: входные числа, зелёный: ссылка на лист «Данные», чёрный: расчёты, жёлтая заливка: фильтры.", NOTE),
    ]
    for i, (t, f) in enumerate(text, 1):
        c = wi.cell(i, 1, t)
        c.font, c.alignment = f, Alignment(wrap_text=True, vertical="top")
    st = prepare_staff(e)
    add_gross(wb, nd, st, gross_profit_1c())
    wb["Мост"].title = "Мост прямой маржи"
    wb["Срезы"].title = "Срезы прямой маржи"
    return wb, e, st


# ------------------------------------------------------------------ достройка до валовой маржи
UNALLOC = "—"   # код строк уровня клиники: медперсонал и амортизация в плане по направлениям не заданы
CONTRIB_PLAN = 0.22   # допущение демо: плоская эффективная ставка взносов, как в генераторе данных


def prepare_staff(e):
    """Входные данные для достройки: оплата врачей, медперсонал, амортизация (план и факт)."""
    pay = pd.read_csv(CORR / "payroll_1c.csv")
    doctors = pd.read_csv(CORR / "doctors.csv")
    opex = pd.read_csv(CORR / "budget_opex.csv")
    post = pd.read_csv(CORR / "postings_1c.csv", dtype={"debit": str, "credit": str})

    cell = e.groupby(CELL, as_index=False).agg(
        plan_rev=("q_p_pp", "sum"), fact_rev_pp=("q_f_pp", "sum"), fact_rev=("fact_revenue", "sum"))
    pct = doctors.groupby(["clinic_id", "direction_code"], as_index=False).pay_pct.mean().rename(columns={"pay_pct": "plan_pct"})
    doc = (pay[pay.role == "doctor"].groupby(CELL, as_index=False)
           .agg(doc_gross=("gross", "sum"), doc_contrib=("contributions", "sum")))
    cell = cell.merge(pct, on=["clinic_id", "direction_code"], how="left").merge(doc, on=CELL, how="left")
    if cell.plan_pct.isna().any():
        raise ValueError("Нет плановой доли врачей для части ячеек")
    cell[["doc_gross", "doc_contrib"]] = cell[["doc_gross", "doc_contrib"]].fillna(0)
    for c in ("nurse_plan", "nurse_gross", "nurse_contrib", "depr_plan", "depr_fact"):
        cell[c] = 0.0

    nur = (pay[pay.role == "nurse"].groupby(["period", "clinic_id"], as_index=False)
           .agg(nurse_gross=("gross", "sum"), nurse_contrib=("contributions", "sum")))
    npl = opex[opex.article == "Оплата труда медперсонала"][["period", "clinic_id", "plan_amount"]].rename(columns={"plan_amount": "nurse_plan"})
    dpl = opex[opex.article == "Амортизация"][["period", "clinic_id", "plan_amount"]].rename(columns={"plan_amount": "depr_plan"})
    dep = post[(post.doc_type == "Амортизация ОС") & (post.debit == "20")].copy()
    dep["period"] = dep.date.str[:7]
    dep = dep.groupby(["period", "clinic_id"], as_index=False).amount.sum().rename(columns={"amount": "depr_fact"})
    cl = npl.merge(dpl, on=["period", "clinic_id"]).merge(nur, on=["period", "clinic_id"], how="left").merge(
        dep, on=["period", "clinic_id"], how="left").fillna(0)
    cl["direction_code"] = UNALLOC
    for c in ("plan_rev", "fact_rev_pp", "fact_rev", "plan_pct", "doc_gross", "doc_contrib"):
        cl[c] = 0.0

    cols = CELL + ["plan_rev", "fact_rev_pp", "fact_rev", "plan_pct", "doc_gross", "doc_contrib",
                   "nurse_plan", "nurse_gross", "nurse_contrib", "depr_plan", "depr_fact"]
    st = pd.concat([cell[cols], cl[cols]], ignore_index=True)
    st["_o"] = (st.direction_code == UNALLOC).astype(int)
    st = st.sort_values(["period", "clinic_id", "_o", "direction_code"]).drop(columns="_o").reset_index(drop=True)
    return st.round({c: 2 for c in cols[3:] if c != "plan_pct"})


def staff_effects(st, rate=CONTRIB_PLAN):
    """То же разложение в pandas: контроль формул Excel."""
    s = st.copy()
    k = s.plan_pct * (1 + rate)
    s["plan_doc_c"] = s.plan_rev * k
    s["doc_vol"] = -(s.fact_rev_pp - s.plan_rev) * k
    s["doc_price"] = -(s.fact_rev - s.fact_rev_pp) * k
    s["doc_grade"] = s.fact_rev * k - (s.doc_gross + s.doc_contrib)
    s["nurse_eff"] = s.nurse_plan * (1 + rate) - (s.nurse_gross + s.nurse_contrib)
    s["depr_eff"] = s.depr_plan - s.depr_fact
    s["plan_costs"] = s.plan_doc_c + s.nurse_plan * (1 + rate) + s.depr_plan
    s["fact_costs"] = s.doc_gross + s.doc_contrib + s.nurse_gross + s.nurse_contrib + s.depr_fact
    s["check"] = s.plan_costs - s.fact_costs - (s.doc_vol + s.doc_price + s.doc_grade + s.nurse_eff + s.depr_eff)
    return s


GROSS_FACTORS = [
    ("Врачи: объём (гибкий план)", "P",
     "Недобор объёма снижает оплату врачей (это процент от выручки). Не экономия, а следствие недобора"),
    ("Врачи: цена и скидки (гибкий план)", "Q",
     "Скидки и другая индексация снижают выручку, а вместе с ней оплату врачей"),
    ("Врачи: структура грейдов и доля", "R",
     "Фактическая доля оплаты врачей против плановой на фактической выручке: реальная эффективность затрат"),
    ("Медперсонал", "S", "Фиксированная статья, на уровне клиники (в плане по направлениям не задана)"),
    ("Амортизация", "T", "Фиксированная статья, на уровне клиники (в плане по направлениям не задана)"),
]


def add_gross(wb, nd, st, gross_1c):
    """Листы «Данные ФОТ», «Мост валовая», «Срезы валовая»."""
    n = len(st)
    last = n + 1
    wf = wb.create_sheet("Данные ФОТ")
    hdr = ["Период", "Клиника", "Направление", "Выручка, план", "Факт кол-во × план цена", "Выручка, факт",
           "Доля врачей, план", "Оплата врачей, факт", "Взносы с оплаты врачей, факт",
           "Медперсонал, план", "Медперсонал, факт", "Взносы с медперсонала, факт",
           "Амортизация, план", "Амортизация, факт",
           "Врачи с взносами, план", "Врачи: объём", "Врачи: цена и скидки", "Врачи: грейды и доля",
           "Медперсонал", "Амортизация", "Затраты, план", "Затраты, факт", "Проверка"]
    for i, h in enumerate(hdr, 1):
        wf.cell(1, i, h)
    head(wf, 1, 1, len(hdr))
    wf.row_dimensions[1].height = 42
    wf["Y1"], wf["Y1"].font = "Ставка взносов в плане (допущение демо):", BOLD
    wf["Z1"], wf["Z1"].font, wf["Z1"].fill = CONTRIB_PLAN, BLUE, INPUT_FILL
    wf["Z1"].number_format = "0%"
    wf["Y2"] = "Плоская эффективная ставка, без предельной базы. Строки с направлением «—»: уровень клиники."
    wf["Y2"].font = NOTE
    for r, row in enumerate(st.itertuples(index=False), 2):
        vals = list(row)
        for c, v in enumerate(vals, 1):
            cell = wf.cell(r, c, v)
            cell.font = BLUE if c >= 4 else BLACK
            if c >= 4:
                cell.number_format = "0.0000" if c == 7 else "#,##0.00"
        k = f"G{r}*(1+$Z$1)"
        wf.cell(r, 15, f"=D{r}*{k}")
        wf.cell(r, 16, f"=-(E{r}-D{r})*{k}")
        wf.cell(r, 17, f"=-(F{r}-E{r})*{k}")
        wf.cell(r, 18, f"=F{r}*{k}-(H{r}+I{r})")
        wf.cell(r, 19, f"=J{r}*(1+$Z$1)-(K{r}+L{r})")
        wf.cell(r, 20, f"=M{r}-N{r}")
        wf.cell(r, 21, f"=O{r}+J{r}*(1+$Z$1)+M{r}")
        wf.cell(r, 22, f"=H{r}+I{r}+K{r}+L{r}+N{r}")
        wf.cell(r, 23, f"=U{r}-V{r}-(P{r}+Q{r}+R{r}+S{r}+T{r})")
        for c in range(15, 24):
            wf.cell(r, c).font = BLACK
            wf.cell(r, c).number_format = "#,##0.00"
    for i, w in enumerate([9, 8, 11] + [14] * 20, 1):
        wf.column_dimensions[col(i)].width = w
    wf.column_dimensions["Y"].width = 40
    wf.freeze_panes = "D2"
    wf.auto_filter.ref = f"A1:W{last}"
    Z = lambda c: f"'Данные ФОТ'!${c}$2:${c}${last}"   # noqa: E731
    D = lambda c: f"'Данные'!${c}$2:${c}${nd}"          # noqa: E731

    def filt(ws, row_clinic=2, row_dir=3):
        for ref, wild, vals in ((f"B{row_clinic}", f"C{row_clinic}", CLINICS), (f"B{row_dir}", f"C{row_dir}", list(DIRS))):
            ws[ref], ws[ref].fill, ws[ref].font = "Все", INPUT_FILL, BLUE
            ws[wild] = f'=IF($B${ref[1:]}="Все","*",$B${ref[1:]})'
            ws[wild].font = Font(name=F, size=8, color="808080")
            dv = DataValidation(type="list", formula1='"' + ",".join(["Все"] + vals) + '"', allow_blank=False)
            ws.add_data_validation(dv)
            dv.add(ref)

    def S(rng_fn, col_):  # noqa: E306
        return f"SUMIFS({rng_fn(col_)},{rng_fn('B')},$C$2,{rng_fn('C')},$C$3)"

    # ---------------------------------------------------------------- Мост валовая
    ws = wb.create_sheet("Мост валовая", 1)
    ws["A1"], ws["A1"].font = "Факторный анализ валовой маржи, 2025 (тыс. ₽)", TITLE
    ws["A2"], ws["A2"].font = "Клиника (фильтр):", BOLD
    ws["A3"], ws["A3"].font = "Направление (фильтр):", BOLD
    filt(ws)
    ws["D2"], ws["D3"] = "← Все, C01, C02 или C03", "← Все, DENT, FAM, COSM, PLAS или LASER"
    ws["D2"].font = ws["D3"].font = NOTE
    ws["A4"] = ("Плюс = выгодно для маржи, минус = невыгодно. Валовая маржа = выручка − материалы − лаборатории − оплата труда врачей "\
                "и медперсонала с взносами − амортизация. Взносы: допущение демо (плоские 22%).")
    ws["A4"].font = NOTE
    ws["A5"] = ("Оплата врачей считается на гибком плане: плановая доля × фактическая выручка. При выбранном направлении медперсонал и "\
                "амортизация не входят (в плане по направлениям не заданы), маржа называется «после оплаты врачей».")
    ws["A5"].font = NOTE
    for i, h in enumerate(["Показатель", "Сумма", "% плановой маржи", "Что означает"], 1):
        ws.cell(7, i, h)
    head(ws, 7, 1, 8)
    ws["D7"].alignment = Alignment(horizontal="left", vertical="center")
    meaning = {
        "Объём": "Больше или меньше услуг, чем в плане, при плановых ценах и составе (за вычетом материалов и лабораторий)",
        "Структура услуг": "Состав услуг внутри направления сместился к более или менее маржинальным",
        "Цена: прайс (индексация)": "Действовавший прайс отличается от плановой цены (другая индексация)",
        "Цена: скидки": "Скидки пациентам снизили факт. цену относительно прайса",
        "Материалы": "Материалы на единицу услуги дороже или дешевле плана",
        "Лаборатории и внешние услуги": "Лаборатории и внешние услуги на единицу дороже или дешевле плана"}
    lab_plan = '=IF($B$3="Все","Плановая валовая маржа","Плановая маржа после оплаты врачей")'
    lab_fact = '=IF($B$3="Все","Фактическая валовая маржа","Фактическая маржа после оплаты врачей")'
    ws.cell(8, 1, lab_plan).font = BOLD
    ws.cell(8, 2, f"={S(D, 'V')}-{S(Z, 'U')}").font = BOLD
    r = 9
    for name, c in FACTORS:
        ws.cell(r, 1, name).font = BLACK
        ws.cell(r, 2, f"={S(D, c)}").font = GREEN_LINK
        ws.cell(r, 3, f"=IF($B$8=0,0,B{r}/$B$8)")
        ws.cell(r, 4, meaning[name]).font = NOTE
        r += 1
    for name, c, txt in GROSS_FACTORS:
        ws.cell(r, 1, name).font = BLACK
        ws.cell(r, 2, f"={S(Z, c)}").font = GREEN_LINK
        ws.cell(r, 3, f"=IF($B$8=0,0,B{r}/$B$8)")
        ws.cell(r, 4, txt).font = NOTE
        r += 1
    fr = r
    ws.cell(fr, 1, lab_fact).font = BOLD
    ws.cell(fr, 2, f"=B8+SUM(B9:B{fr - 1})").font = BOLD
    ws.cell(fr, 3, f"=IF($B$8=0,0,B{fr}/$B$8-1)")
    ws.cell(fr + 1, 1, "Контроль: факт маржи по данным минус факт по мосту").font = NOTE
    ws.cell(fr + 1, 2, f"=({S(D, 'AC')}-{S(Z, 'V')})-B{fr}").font = NOTE
    ws.cell(fr + 1, 4, "должно быть 0").font = NOTE
    for rr in range(8, fr + 2):
        ws.cell(rr, 2).number_format = NUM if rr != fr + 1 else "#,##0.00"
        ws.cell(rr, 3).number_format = PCT
    for c in range(1, 5):
        ws.cell(fr, c).border = TOP
        ws.cell(8, c).fill = SEC_FILL
    sign_format(ws, f"B9:C{fr - 1}")
    ws.cell(fr, 3).font = BOLD

    q0 = fr + 3
    for i, h in enumerate(["", "План", "Факт", "Откл."], 1):
        ws.cell(q0, i, h)
    head(ws, q0, 1, 4)
    ws.cell(q0 + 1, 1, "Выручка")
    ws.cell(q0 + 1, 2, f"={S(D, 'Q')}")
    ws.cell(q0 + 1, 3, f"={S(D, 'J')}")
    ws.cell(q0 + 2, 1, "Валовая маржа (или маржа после оплаты врачей)")
    ws.cell(q0 + 2, 2, "=B8")
    ws.cell(q0 + 2, 3, f"=B{fr}")
    ws.cell(q0 + 3, 1, "Маржа, % выручки")
    for c in (2, 3):
        L = col(c)
        ws.cell(q0 + 3, c, f"=IF({L}{q0 + 1}=0,0,{L}{q0 + 2}/{L}{q0 + 1})")
    ws.cell(q0 + 1, 4, f"=C{q0 + 1}-B{q0 + 1}")
    ws.cell(q0 + 2, 4, f"=C{q0 + 2}-B{q0 + 2}")
    ws.cell(q0 + 3, 4, f"=(C{q0 + 3}-B{q0 + 3})*100")
    for rr in (q0 + 1, q0 + 2):
        for c in (2, 3, 4):
            ws.cell(rr, c).number_format = NUM
    for c in (2, 3):
        ws.cell(q0 + 3, c).number_format = PCT
    ws.cell(q0 + 3, 4).number_format = PP
    for rr in range(q0 + 1, q0 + 4):
        for c in range(1, 5):
            ws.cell(rr, c).font = BLACK
    sign_format(ws, f"D{q0 + 1}:D{q0 + 3}")

    k0 = q0 + 5
    ws.cell(k0, 1, "Сверка с 1С (вся сеть, фильтры не действуют)").font = BOLD
    for i, h in enumerate(["", "Сумма"], 1):
        ws.cell(k0 + 1, i, h)
    head(ws, k0 + 1, 1, 2)
    ws.cell(k0 + 2, 1, "Валовая прибыль по 1С (90.01 минус обороты счёта 20), значение из проводок")
    ws.cell(k0 + 2, 2, gross_1c).font = BLUE
    ws.cell(k0 + 3, 1, "Валовая прибыль по книге (факт прямой маржи минус ФОТ с взносами и амортизация)")
    ws.cell(k0 + 3, 2, f"=SUM({D('AC')})-SUM({Z('V')})")
    ws.cell(k0 + 4, 1, "Разница, ₽ (выручка ERP и 1С после исправлений расходится на ≈100 ₽ в пределах допуска сверки)")
    ws.cell(k0 + 4, 2, f"=B{k0 + 3}-B{k0 + 2}")
    ws.cell(k0 + 5, 1, "Статус")
    ws.cell(k0 + 5, 2, f'=IF(ABS(B{k0 + 4})<=100,"СХОДИТСЯ","РАСХОДИТСЯ")').font = BOLD
    for rr in range(k0 + 2, k0 + 5):
        ws.cell(rr, 2).number_format = "#,##0"
        if rr != k0 + 2:
            ws.cell(rr, 2).font = BLACK
        ws.cell(rr, 1).font = BLACK
    ws.cell(k0 + 5, 1).font = BLACK
    ws.column_dimensions["A"].width = 58
    for c_ in "BCDEFGH":
        ws.column_dimensions[c_].width = 14

    ch = BarChart()
    ch.type, ch.title = "bar", "Влияние факторов на валовую маржу, тыс. ₽"
    ch.add_data(Reference(ws, min_col=2, min_row=9, max_row=fr - 1), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=1, min_row=9, max_row=fr - 1))
    ch.legend = None
    ch.series[0].invertIfNegative = False
    ch.series[0].graphicalProperties.solidFill = "2E75B6"
    ch.y_axis.numFmt = "#,##0,"
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    ch.x_axis.scaling.orientation = "maxMin"
    ch.x_axis.tickLblPos = "low"
    ch.height, ch.width = 10, 24
    ws.add_chart(ch, f"A{k0 + 8}")

    # ---------------------------------------------------------------- Срезы валовая
    wc = wb.create_sheet("Срезы валовая", 2)
    wc["A1"], wc["A1"].font = "Факторы валовой маржи по направлениям и клиникам, 2025 (тыс. ₽)", TITLE
    wc["A2"] = ("Строка «Не распределено» это медперсонал и амортизация: в плане они заданы по клиникам, без направлений. "\
                "Оплата врачей показана на гибком плане; «грейды и доля» это отклонение фактической доли от плановой.")
    wc["A2"].font = NOTE
    heads = ["План ВМ", "Объём", "Структура услуг", "Цена: прайс", "Цена: скидки", "Материалы", "Лаборатории",
             "Врачи: объём", "Врачи: цена", "Врачи: грейды и доля", "Медперсонал", "Амортизация",
             "Факт ВМ", "Откл. ВМ", "ВМ план, %", "ВМ факт, %", "Откл., п.п."]
    srcs = [("D", "W"), ("D", "X"), ("D", "Y"), ("D", "Z"), ("D", "AA"), ("D", "AB"),
            ("Z", "P"), ("Z", "Q"), ("Z", "R"), ("Z", "S"), ("Z", "T")]
    KEYCOL = 21  # U

    def slice_table(top, title, key_col, keys, names, filter_label, filter_ref, filter_wild, vals, filter_col):
        wc.cell(top, 1, title).font = BOLD
        wc.cell(top + 1, 1, filter_label).font = BOLD
        wc[filter_ref], wc[filter_ref].fill, wc[filter_ref].font = "Все", INPUT_FILL, BLUE
        fr_ = f"${filter_ref[0]}${filter_ref[1:]}"
        fw = f"${filter_wild[0]}${filter_wild[1:]}"
        wc[filter_wild] = f'=IF({fr_}="Все","*",{fr_})'
        wc[filter_wild].font = Font(name=F, size=8, color="808080")
        dv = DataValidation(type="list", formula1='"' + ",".join(["Все"] + vals) + '"', allow_blank=False)
        wc.add_data_validation(dv)
        dv.add(filter_ref)
        h = top + 3
        wc.cell(h, 1, "Срез")
        for i, t in enumerate(heads):
            wc.cell(h, 2 + i, t)
        head(wc, h, 1, 1 + len(heads))
        wc.row_dimensions[h].height = 30
        allk = keys + ["*"]
        for j, (k, nm) in enumerate(zip(allk, names + ["Итого"])):
            rr = h + 1 + j
            bold = k == "*"
            fn = BOLD if bold else BLACK
            wc.cell(rr, 1, nm).font = fn
            wc.cell(rr, KEYCOL, k).font = Font(name=F, size=8, color="808080")
            kc = f"$U{rr}"
            sm = lambda sh, c: (f"SUMIFS({D(c) if sh == 'D' else Z(c)},"      # noqa: E731
                                f"{D(key_col) if sh == 'D' else Z(key_col)},{kc},"
                                f"{D(filter_col) if sh == 'D' else Z(filter_col)},{fw})")
            wc.cell(rr, 2, f"={sm('D', 'V')}-{sm('Z', 'U')}")
            for i, (sh, c) in enumerate(srcs):
                wc.cell(rr, 3 + i, f"={sm(sh, c)}")
            wc.cell(rr, 14, f"=B{rr}+SUM(C{rr}:M{rr})")
            wc.cell(rr, 15, f"=N{rr}-B{rr}")
            wc.cell(rr, 16, f"=IF({sm('D', 'Q')}=0,0,B{rr}/{sm('D', 'Q')})")
            wc.cell(rr, 17, f"=IF({sm('D', 'J')}=0,0,N{rr}/{sm('D', 'J')})")
            wc.cell(rr, 18, f"=(Q{rr}-P{rr})*100")
            for c in range(2, 16):
                wc.cell(rr, c).number_format = NUM
            wc.cell(rr, 16).number_format = wc.cell(rr, 17).number_format = PCT
            wc.cell(rr, 18).number_format = PP
            for c in range(2, 19):
                wc.cell(rr, c).font = fn
            if bold:
                for c in range(1, 19):
                    wc.cell(rr, c).border = TOP
        lastr = h + len(allk)
        sign_format(wc, f"C{h + 1}:M{lastr}")
        sign_format(wc, f"O{h + 1}:O{lastr}")
        sign_format(wc, f"R{h + 1}:R{lastr}")
        return lastr

    dkeys = list(DIRS) + [UNALLOC]
    dnames = list(DIRS.values()) + ["Не распределено (медперсонал, амортизация)"]
    last_ = slice_table(4, "По направлениям", "C", dkeys, dnames, "Клиника (фильтр):", "B5", "C5", CLINICS, "B")
    slice_table(last_ + 3, "По клиникам", "B", CLINICS, ["Невский", "Север", "Юг"],
                "Направление (фильтр):", f"B{last_ + 4}", f"C{last_ + 4}", list(DIRS), "C")
    wc.column_dimensions["A"].width = 44
    for c in range(2, 19):
        wc.column_dimensions[col(c)].width = 13
    wc.column_dimensions["U"].width = 6
    wc.freeze_panes = "B4"
    return wb



def main():
    OUT.mkdir(parents=True, exist_ok=True)
    d, mat1c = prepare()
    wb, e, st = build(d)
    path = OUT / "FactorAnalysis_2025.xlsx"
    wb.save(path)
    recalculate(path)

    # контроль в pandas
    t = e[["plan_margin", "volume", "mix", "price_list", "discount", "materials", "labs", "fact_margin"]].sum()
    print(f"Сохранено: {path}  (строк: {len(d)})")
    print(f"Макс. остаток разложения по строкам: {e.check.abs().max():.6f}")
    print(f"Затраты факта: материалы {d.mat_fact.sum():,.0f} (1С {mat1c.material_fact.sum():,.0f}); "
          f"лаборатории {d.ext_fact.sum():,.0f} (1С {mat1c.external_fact.sum():,.0f})")
    print(f"Плановая прямая маржа {t.plan_margin / 1e6:,.2f} млн ₽ -> факт {t.fact_margin / 1e6:,.2f} млн ₽ "
          f"(откл. {(t.fact_margin - t.plan_margin) / 1e6:,.2f})")
    for k, nm in (("volume", "Объём"), ("mix", "Структура услуг"), ("price_list", "Цена: прайс"),
                  ("discount", "Цена: скидки"), ("materials", "Материалы"), ("labs", "Лаборатории")):
        print(f"  {nm:<18}{t[k] / 1e6:>8.2f} млн ₽")

    # достройка до валовой маржи
    se = staff_effects(st)
    tot = se[["plan_costs", "fact_costs", "doc_vol", "doc_price", "doc_grade", "nurse_eff", "depr_eff"]].sum()
    gm_plan = t.plan_margin - tot.plan_costs
    gm_fact = t.fact_margin - tot.fact_costs
    print(f"Валовая маржа: план {gm_plan / 1e6:,.2f} -> факт {gm_fact / 1e6:,.2f} млн ₽ (откл. {(gm_fact - gm_plan) / 1e6:,.2f})")
    for k, nm in (("doc_vol", "Врачи: объём"), ("doc_price", "Врачи: цена и скидки"), ("doc_grade", "Врачи: грейды и доля"),
                  ("nurse_eff", "Медперсонал"), ("depr_eff", "Амортизация")):
        print(f"  {nm:<22}{tot[k] / 1e6:>8.2f} млн ₽")
    g1c = gross_profit_1c()
    print(f"Макс. остаток по строкам ФОТ: {se.check.abs().max():.6f}")
    print(f"Валовая прибыль: книга {gm_fact:,.0f} / 1С {g1c:,.0f} / разница {gm_fact - g1c:,.0f} ₽")


if __name__ == "__main__":
    main()

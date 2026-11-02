"""
Шаг 8: отчёт о прибылях и убытках (P&L) за 2025 год в Excel.

Вход:  data/corrected/ (проводки и ОСВ после исправлений сверки)
Выход: data/reports/PnL_2025.xlsx

Книга живая: на листе «Данные» лежат агрегированные проводки 1С (месяц, клиника,
направление, статья), а все отчётные листы считаются формулами SUMIFS. Меняешь
клинику или направление в жёлтых ячейках листа «P&L», и отчёт пересчитывается.

Листы:
  Описание            - источники, правила расчёта, ограничения, статус сверки
  P&L                 - помесячный P&L с фильтром по клинике и направлению + график
  По клиникам         - P&L за год по клиникам
  По направлениям     - валовая маржа по направлениям (косвенные расходы не распределяются)
  Сверка с ОСВ        - контроль: P&L из проводок против оборотов ОСВ по счетам 90 и 99
  Данные              - агрегированные проводки (источник для формул)

Правила:
  Выручка            = кредитовые обороты 90.01 по документам «Реализация услуг»
  Себестоимость      = дебетовые обороты счёта 20 (врачи, медперсонал, взносы,
                       материалы, лаборатории, амортизация)
  Управленческие     = дебетовые обороты счёта 26 (администрация, аренда, коммунальные, эквайринг)
  Коммерческие       = дебетовые обороты счёта 44 (маркетинг)
  Закрытие месяца (счета 90 -> 99) в P&L не участвует.
Налог на прибыль и НДС вне рамок демо: P&L заканчивается прибылью до налога.

Запуск:  python src/reports/build_pnl.py
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.recalc import recalculate  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "reports"
STR = {"debit": str, "credit": str, "account": str}

DIRS = {"DENT": "Стоматология", "FAM": "Семейная медицина", "COSM": "Косметология",
        "PLAS": "Пластическая хирургия", "LASER": "Лазерная коррекция зрения",
        "ОБЩ": "Общеклинические (не распределены)"}
ARTICLE = {
    "Начисление ЗП врачей": "Оплата труда врачей",
    "Начисление ЗП медперсонала": "Оплата труда медперсонала",
    "Списание материалов": "Материалы",
    "Услуги лабораторий и сторонних организаций": "Лаборатории и внешние услуги",
    "Амортизация ОС": "Амортизация",
    "Начисление ЗП административного персонала": "Оплата труда административного персонала",
    "Аренда помещений": "Аренда",
    "Коммунальные услуги": "Коммунальные услуги",
    "Комиссия эквайринга": "Комиссия эквайринга",
    "Рекламные услуги": "Маркетинг",
}
BLOCK = {"20": "Себестоимость", "26": "Управленческие", "44": "Коммерческие"}
COST_ARTICLES = ["Оплата труда врачей", "Оплата труда медперсонала", "Страховые взносы (себестоимость)",
                 "Материалы", "Лаборатории и внешние услуги", "Амортизация"]
MGMT_ARTICLES = ["Оплата труда административного персонала", "Страховые взносы (администрация)",
                 "Аренда", "Коммунальные услуги", "Комиссия эквайринга"]

# --- стили ---
F = "Arial"
f_norm = Font(name=F, size=10)
f_bold = Font(name=F, size=10, bold=True)
f_head = Font(name=F, size=10, bold=True, color="FFFFFF")
f_title = Font(name=F, size=14, bold=True)
f_link = Font(name=F, size=10, color="008000")          # ссылка на другой лист
f_input = Font(name=F, size=10, color="0000FF")         # ввод / жёстко заданное значение
f_note = Font(name=F, size=9, italic=True, color="595959")
fill_head = PatternFill("solid", fgColor="1F3864")
fill_sub = PatternFill("solid", fgColor="D9E1F2")
fill_key = PatternFill("solid", fgColor="FFFF00")
thin = Side(style="thin", color="BFBFBF")
top_border = Border(top=Side(style="thin", color="000000"))
NUM = "#,##0,;(#,##0,);-"          # рубли, отображаются в тысячах
PCT = "0.0%;(0.0%);-"


def prepare():
    post = pd.read_csv(CORR / "postings_1c.csv", dtype=STR)
    osv = pd.read_csv(CORR / "osv_1c.csv", dtype=STR | {c: float for c in [
        "opening_debit", "opening_credit", "turnover_debit", "turnover_credit",
        "closing_debit", "closing_credit"]})
    oper = post[post.doc_type != "Закрытие месяца"].copy()
    oper["period"] = oper.date.str[:7]
    oper["direction_code"] = oper.direction_code.fillna("ОБЩ")

    rev = oper[(oper.credit == "90.01") & (oper.doc_type == "Реализация услуг")].copy()
    rev["block"], rev["article"] = "Выручка", "Выручка"

    cst = oper[oper.debit.isin(BLOCK)].copy()
    cst["block"] = cst.debit.map(BLOCK)

    def art(r):
        if r.doc_type == "Начисление страховых взносов":
            return "Страховые взносы (себестоимость)" if r.debit == "20" else "Страховые взносы (администрация)"
        return ARTICLE[r.doc_type]          # KeyError = неизвестный тип документа, пусть падает
    cst["article"] = cst.apply(art, axis=1)

    data = pd.concat([rev, cst]).groupby(
        ["period", "clinic_id", "direction_code", "block", "article"]).amount.sum().round(2).reset_index()
    data["direction"] = data.direction_code.map(DIRS)
    data = data[["period", "clinic_id", "direction_code", "direction", "block", "article", "amount"]]
    data = data.sort_values(["period", "clinic_id", "direction_code", "block", "article"]).reset_index(drop=True)

    o = osv[osv.account.isin(["90.01", "90.02", "90.07", "90.08", "99"])]
    o = o.groupby(["period", "account"])[["turnover_debit", "turnover_credit"]].sum()
    return data, o


def head(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.font, cell.fill = f_head, fill_head
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build(data, osv_turn):
    wb = Workbook()
    n = len(data) + 1                     # последняя строка данных
    periods = [f"2025-{m:02d}" for m in range(1, 13)]
    R = lambda c: f"'Данные'!${c}$2:${c}${n}"

    # ---------------------------------------------------------------- Данные
    wd = wb.active
    wd.title = "Данные"
    cols = ["Период", "Клиника", "Код направления", "Направление", "Блок P&L", "Статья", "Сумма, ₽"]
    for i, h in enumerate(cols, 1):
        wd.cell(1, i, h)
    head(wd, 1, 1, 7)
    for r, row in enumerate(data.itertuples(index=False), 2):
        for c, v in enumerate(row, 1):
            cell = wd.cell(r, c, v)
            cell.font = f_input if c == 7 else f_norm
        wd.cell(r, 7).number_format = "#,##0.00"
    for c, w in zip("ABCDEFG", [10, 9, 16, 34, 16, 42, 16]):
        wd.column_dimensions[c].width = w
    wd.freeze_panes = "A2"
    wd.auto_filter.ref = f"A1:G{n}"

    # ---------------------------------------------------------------- P&L
    ws = wb.create_sheet("P&L", 0)
    ws["A1"], ws["A1"].font = "P&L сети клиник, 2025 (тыс. ₽)", f_title
    ws["A2"], ws["A3"] = "Клиника (фильтр):", "Направление (фильтр):"
    ws["B2"], ws["B3"] = "Все", "Все"
    for c in ("B2", "B3"):
        ws[c].font, ws[c].fill = f_input, fill_key
    ws["C2"] = '=IF($B$2="Все","*",$B$2)'
    ws["C3"] = '=IF($B$3="Все","*",$B$3)'
    ws["D2"] = "← выбери из списка (Все / C01 / C02 / C03)"
    ws["D3"] = "← Все / DENT / FAM / COSM / PLAS / LASER. При выборе направления косвенные расходы не распределены"
    for c in ("C2", "C3", "D2", "D3"):
        ws[c].font = f_note
    dv1 = DataValidation(type="list", formula1='"Все,C01,C02,C03"', allow_blank=False)
    dv2 = DataValidation(type="list", formula1='"Все,DENT,FAM,COSM,PLAS,LASER"', allow_blank=False)
    ws.add_data_validation(dv1)
    ws.add_data_validation(dv2)
    dv1.add("B2")
    dv2.add("B3")

    ws.cell(5, 1, "Статья")
    for i, p in enumerate(periods):
        ws.cell(5, 2 + i, p)
    ws.cell(5, 14, "Итого 2025")
    ws.cell(5, 15, "% выручки")
    head(ws, 5, 1, 15)

    layout = {}   # название строки -> номер
    r = 6

    def sumifs_row(label, article, bold=False):
        nonlocal r
        layout[label] = r
        ws.cell(r, 1, label).font = f_bold if bold else f_norm
        for i in range(12):
            L = col(2 + i)
            ws.cell(r, 2 + i, f"=SUMIFS({R('G')},{R('A')},{L}$5,{R('F')},$A{r},{R('B')},$C$2,{R('C')},$C$3)")
            ws.cell(r, 2 + i).font = f_link
        r += 1

    def calc_row(label, formula_fn, bold=True, pct=False, border=False):
        nonlocal r
        layout[label] = r
        ws.cell(r, 1, label).font = f_bold if bold else f_norm
        for i in range(13):  # 12 месяцев + итог
            L = col(2 + i)
            cell = ws.cell(r, 2 + i, formula_fn(L))
            cell.font = f_bold if bold else f_norm
            cell.number_format = PCT if pct else NUM
            if border:
                cell.border = top_border
        r += 1

    def section(label):
        nonlocal r
        c = ws.cell(r, 1, label)
        c.font = f_bold
        for k in range(1, 16):
            ws.cell(r, k).fill = fill_sub
        r += 1

    sumifs_row("Выручка", "Выручка", bold=True)
    section("Себестоимость услуг")
    first_cost = r
    for a in COST_ARTICLES:
        sumifs_row(a, a)
    last_cost = r - 1
    calc_row("Итого себестоимость", lambda L: f"=SUM({L}{first_cost}:{L}{last_cost})", border=True)
    calc_row("Валовая прибыль", lambda L: f"={L}{layout['Выручка']}-{L}{layout['Итого себестоимость']}")
    calc_row("Валовая маржа, %", lambda L: f"=IF({L}{layout['Выручка']}=0,0,{L}{layout['Валовая прибыль']}/{L}{layout['Выручка']})",
             bold=False, pct=True)
    section("Управленческие расходы")
    first_m = r
    for a in MGMT_ARTICLES:
        sumifs_row(a, a)
    last_m = r - 1
    calc_row("Итого управленческие", lambda L: f"=SUM({L}{first_m}:{L}{last_m})", border=True)
    section("Коммерческие расходы")
    sumifs_row("Маркетинг", "Маркетинг")
    calc_row("Прибыль до налога", lambda L: f"={L}{layout['Валовая прибыль']}-{L}{layout['Итого управленческие']}-{L}{layout['Маркетинг']}")
    calc_row("Маржа прибыли, %", lambda L: f"=IF({L}{layout['Выручка']}=0,0,{L}{layout['Прибыль до налога']}/{L}{layout['Выручка']})",
             bold=False, pct=True)
    calc_row("EBITDA", lambda L: f"={L}{layout['Прибыль до налога']}+{L}{layout['Амортизация']}")
    calc_row("EBITDA, %", lambda L: f"=IF({L}{layout['Выручка']}=0,0,{L}{layout['EBITDA']}/{L}{layout['Выручка']})",
             bold=False, pct=True)
    last_row = r - 1

    # итоги по году и % выручки для всех строк с суммами
    pct_rows = {"Валовая маржа, %", "Маржа прибыли, %", "EBITDA, %"}
    for label, rr in layout.items():
        if label in pct_rows:
            ws.cell(rr, 14, f"=IF(N{layout['Выручка']}=0,0,N{layout[label.replace(', %', '').replace('Валовая маржа', 'Валовая прибыль').replace('Маржа прибыли', 'Прибыль до налога')]}/N{layout['Выручка']})")
            ws.cell(rr, 14).number_format, ws.cell(rr, 14).font = PCT, f_norm
            continue
        if ws.cell(rr, 14).value is None:
            ws.cell(rr, 14, f"=SUM(B{rr}:M{rr})")
        ws.cell(rr, 14).number_format = NUM
        ws.cell(rr, 14).font = f_bold if ws.cell(rr, 2).font.bold else f_norm
        ws.cell(rr, 15, f"=IF($N${layout['Выручка']}=0,0,N{rr}/$N${layout['Выручка']})")
        ws.cell(rr, 15).number_format, ws.cell(rr, 15).font = PCT, f_norm
    for rr in range(6, last_row + 1):
        for c in range(2, 14):
            if ws.cell(rr, c).number_format == "General":
                ws.cell(rr, c).number_format = NUM
    # итог по расчётным строкам: формулы уже заданы calc_row для 13-го столбца (N) как для месяцев
    ws.column_dimensions["A"].width = 42
    for c in range(2, 16):
        ws.column_dimensions[col(c)].width = 11.5
    ws.freeze_panes = "B6"
    ws.cell(last_row + 2, 1, "Суммы показаны в тысячах рублей (хранятся в рублях). Зелёный шрифт: значения из листа «Данные». "
                             "Жёлтые ячейки: фильтры.").font = f_note

    ch = BarChart()
    ch.type, ch.title = "col", "Выручка, валовая прибыль и прибыль до налога по месяцам, тыс. ₽"
    for label, color in (("Выручка", "1F3864"), ("Валовая прибыль", "2E75B6"), ("Прибыль до налога", "A9D18E")):
        s = Series(Reference(ws, min_col=2, max_col=13, min_row=layout[label]), title=label)
        s.graphicalProperties.solidFill = color
        ch.series.append(s)
    ch.set_categories(Reference(ws, min_col=2, max_col=13, min_row=5))
    ch.height, ch.width = 9, 24
    ch.y_axis.numFmt = "#,##0,"
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    ws.add_chart(ch, f"A{last_row + 4}")

    # ---------------------------------------------------------------- По клиникам / направлениям
    def static_sheet(title, key_col, keys, names, header_label, name_map=None):
        w = wb.create_sheet(title)
        w["A1"], w["A1"].font = f"{title}: P&L за 2025 год (тыс. ₽)", f_title
        w.cell(3, 1, "Статья")
        for j, k in enumerate(keys):
            w.cell(3, 2 + j, k)
            w.cell(4, 2 + j, names[j]).font = f_note
        tot = 2 + len(keys)
        w.cell(3, tot, "Итого сеть")
        head(w, 3, 1, tot)
        lay, rr = {}, 5

        def s_row(label, bold=False):
            nonlocal rr
            lay[label] = rr
            w.cell(rr, 1, label).font = f_bold if bold else f_norm
            for j in range(len(keys)):
                L = col(2 + j)
                w.cell(rr, 2 + j, f"=SUMIFS({R('G')},{R(key_col)},{L}$3,{R('F')},$A{rr})").font = f_link
                w.cell(rr, 2 + j).number_format = NUM
            rr += 1

        def c_row(label, fn, bold=True, pct=False):
            nonlocal rr
            lay[label] = rr
            w.cell(rr, 1, label).font = f_bold if bold else f_norm
            for j in range(len(keys) + 1):
                L = col(2 + j)
                cell = w.cell(rr, 2 + j, fn(L))
                cell.font = f_bold if bold else f_norm
                cell.number_format = PCT if pct else NUM
            rr += 1

        s_row("Выручка", True)
        c0 = rr
        for a in COST_ARTICLES:
            s_row(a)
        c1 = rr - 1
        c_row("Итого себестоимость", lambda L: f"=SUM({L}{c0}:{L}{c1})")
        c_row("Валовая прибыль", lambda L: f"={L}{lay['Выручка']}-{L}{lay['Итого себестоимость']}")
        c_row("Валовая маржа, %", lambda L: f"=IF({L}{lay['Выручка']}=0,0,{L}{lay['Валовая прибыль']}/{L}{lay['Выручка']})",
              bold=False, pct=True)
        c_row("Доля в выручке сети, %", lambda L: f"=IF(${col(tot)}{lay['Выручка']}=0,0,{L}{lay['Выручка']}/${col(tot)}{lay['Выручка']})",
              bold=False, pct=True)
        if key_col == "B":      # по клиникам косвенные расходы известны
            rr += 1
            m0 = rr
            for a in MGMT_ARTICLES:
                s_row(a)
            m1 = rr - 1
            c_row("Итого управленческие", lambda L: f"=SUM({L}{m0}:{L}{m1})")
            s_row("Маркетинг")
            c_row("Прибыль до налога", lambda L: f"={L}{lay['Валовая прибыль']}-{L}{lay['Итого управленческие']}-{L}{lay['Маркетинг']}")
            c_row("Маржа прибыли, %", lambda L: f"=IF({L}{lay['Выручка']}=0,0,{L}{lay['Прибыль до налога']}/{L}{lay['Выручка']})",
                  bold=False, pct=True)
        # колонка «Итого сеть» для строк с SUMIFS
        for label, row_ in lay.items():
            if w.cell(row_, tot).value is None:
                w.cell(row_, tot, f"=SUM(B{row_}:{col(tot - 1)}{row_})")
                w.cell(row_, tot).number_format = NUM
                w.cell(row_, tot).font = f_bold if label in ("Выручка",) else f_norm
        if key_col == "C":
            w.cell(rr + 1, 1, "Косвенные расходы (аренда, администрация, эквайринг, маркетинг) по направлениям не "
                              "распределяются: смотри лист «По клиникам».").font = f_note
        w.column_dimensions["A"].width = 42
        for c in range(2, tot + 1):
            w.column_dimensions[col(c)].width = 15
        w.freeze_panes = "B5"
        return w

    static_sheet("По клиникам", "B", ["C01", "C02", "C03"], ["Невский", "Север", "Юг"], "Клиника")
    dkeys = ["DENT", "FAM", "COSM", "PLAS", "LASER"]
    static_sheet("По направлениям", "C", dkeys, [DIRS[k] for k in dkeys], "Направление")

    # ---------------------------------------------------------------- Сверка с ОСВ
    wc = wb.create_sheet("Сверка с ОСВ")
    wc["A1"], wc["A1"].font = "Сверка P&L с ОСВ 1С (сеть, ₽)", f_title
    wc["A2"] = ("P&L считается из проводок (лист «Данные»), ОСВ берётся из отдельного файла osv_1c.csv "
                "(обороты счетов 90 и 99). Синие значения: данные из ОСВ, внесены как есть.")
    wc["A2"].font = f_note
    wc.cell(4, 1, "Показатель")
    for i, p in enumerate(periods):
        wc.cell(4, 2 + i, p)
    wc.cell(4, 14, "Итого")
    head(wc, 4, 1, 14)

    def osv_val(p, acc, side):
        try:
            return float(osv_turn.loc[(p, acc), side])
        except KeyError:
            return 0.0

    checks = [
        ("Выручка", "Выручка по P&L", f"=SUMIFS({R('G')},{R('A')},{{L}}$4,{R('E')},\"Выручка\")",
         "Выручка по ОСВ (90.01, кредит)", lambda p: osv_val(p, "90.01", "turnover_credit")),
        ("Себестоимость", "Себестоимость по P&L", f"=SUMIFS({R('G')},{R('A')},{{L}}$4,{R('E')},\"Себестоимость\")",
         "Себестоимость по ОСВ (90.02, дебет)", lambda p: osv_val(p, "90.02", "turnover_debit")),
        ("Управленческие", "Управленческие по P&L", f"=SUMIFS({R('G')},{R('A')},{{L}}$4,{R('E')},\"Управленческие\")",
         "Управленческие по ОСВ (90.08, дебет)", lambda p: osv_val(p, "90.08", "turnover_debit")),
        ("Коммерческие", "Коммерческие по P&L", f"=SUMIFS({R('G')},{R('A')},{{L}}$4,{R('E')},\"Коммерческие\")",
         "Коммерческие по ОСВ (90.07, дебет)", lambda p: osv_val(p, "90.07", "turnover_debit")),
    ]
    r = 5
    diff_rows = []
    pl_rows = {}
    for key, l1, f1, l2, vf in checks:
        wc.cell(r, 1, l1).font = f_norm
        wc.cell(r + 1, 1, l2).font = f_norm
        wc.cell(r + 2, 1, "Разница").font = f_bold
        for i, p in enumerate(periods):
            L = col(2 + i)
            wc.cell(r, 2 + i, f1.replace("{L}", L)).font = f_link
            wc.cell(r + 1, 2 + i, round(vf(p), 2)).font = f_input
            wc.cell(r + 2, 2 + i, f"={L}{r}-{L}{r + 1}").font = f_bold
        for k in range(3):
            wc.cell(r + k, 14, f"=SUM(B{r + k}:M{r + k})").font = f_bold if k == 2 else f_norm
        for k in range(3):
            for c in range(2, 15):
                wc.cell(r + k, c).number_format = "#,##0.00;(#,##0.00);-"
        pl_rows[key] = r
        diff_rows.append(r + 2)
        r += 4
    # прибыль: P&L против оборота 99
    wc.cell(r, 1, "Прибыль до налога по P&L").font = f_norm
    wc.cell(r + 1, 1, "Прибыль по ОСВ (99, кредит минус дебет)").font = f_norm
    wc.cell(r + 2, 1, "Разница").font = f_bold
    for i, p in enumerate(periods):
        L = col(2 + i)
        wc.cell(r, 2 + i, f"={L}{pl_rows['Выручка']}-{L}{pl_rows['Себестоимость']}-{L}{pl_rows['Управленческие']}-{L}{pl_rows['Коммерческие']}")
        wc.cell(r + 1, 2 + i, round(osv_val(p, "99", "turnover_credit") - osv_val(p, "99", "turnover_debit"), 2)).font = f_input
        wc.cell(r + 2, 2 + i, f"={L}{r}-{L}{r + 1}").font = f_bold
    for k in range(3):
        wc.cell(r + k, 14, f"=SUM(B{r + k}:M{r + k})")
        for c in range(2, 15):
            wc.cell(r + k, c).number_format = "#,##0.00;(#,##0.00);-"
            if k < 2 and wc.cell(r + k, c).font.color is None:
                wc.cell(r + k, c).font = f_norm
    diff_rows.append(r + 2)
    status_row = r + 4
    wc.cell(status_row, 1, "Статус сверки (допуск 1 ₽):").font = f_bold
    parts = "+".join(f"SUMPRODUCT(ABS(B{d}:M{d}))" for d in diff_rows)
    wc.cell(status_row, 2, f'=IF({parts}<1,"СХОДИТСЯ","РАСХОЖДЕНИЕ")').font = f_bold
    wc.cell(status_row, 2).fill = fill_key
    wc.column_dimensions["A"].width = 46
    for c in range(2, 15):
        wc.column_dimensions[col(c)].width = 14
    wc.freeze_panes = "B5"

    # ---------------------------------------------------------------- Описание
    wo = wb.create_sheet("Описание", 0)
    wo["A1"], wo["A1"].font = "P&L сети клиник за 2025 год", f_title
    lines = [
        ("Источник", "Проводки и ОСВ 1С после исправлений сверки (папка data/corrected). Данные синтетические."),
        ("Статус сверки с ОСВ", "='Сверка с ОСВ'!B" + str(status_row)),
        ("Единицы", "Хранится в рублях, на листах отображается в тысячах рублей."),
        ("Выручка", "Кредитовые обороты 90.01 по документам «Реализация услуг»."),
        ("Себестоимость", "Дебет счёта 20: оплата труда врачей и медперсонала, страховые взносы, материалы, "
                          "лаборатории и внешние услуги, амортизация."),
        ("Валовая прибыль", "Выручка минус себестоимость (определение сети: прямые расходы направления)."),
        ("Управленческие", "Дебет счёта 26: администрация, взносы администрации, аренда, коммунальные, комиссия эквайринга."),
        ("Коммерческие", "Дебет счёта 44: маркетинг."),
        ("EBITDA", "Прибыль до налога плюс амортизация."),
        ("Закрытие месяца", "Проводки закрытия счетов 90 в 99 в P&L не участвуют."),
        ("Ограничение 1", "Налог на прибыль и НДС не рассчитываются: P&L заканчивается прибылью до налога."),
        ("Ограничение 2", "Два визита без врача (120 700 ₽ выручки) не удалось исправить автоматически; "
                          "они включены в выручку, на расчёт P&L не влияют."),
        ("Ограничение 3", "Косвенные расходы по направлениям не распределяются, поэтому по направлениям "
                          "показана валовая маржа."),
        ("Как пользоваться", "Лист «P&L»: выбери клинику и направление в жёлтых ячейках B2 и B3."),
        ("Цвета", "Синий шрифт: значения из источников. Зелёный: ссылки на лист «Данные». Чёрный: расчёты. "
                  "Жёлтая заливка: ячейки для ввода."),
    ]
    for i, (k, v) in enumerate(lines, 3):
        wo.cell(i, 1, k).font = f_bold
        c = wo.cell(i, 2, v)
        c.font, c.alignment = f_norm, Alignment(wrap_text=True, vertical="top")
        wo.cell(i, 1).alignment = Alignment(vertical="top")
    wo.column_dimensions["A"].width = 24
    wo.column_dimensions["B"].width = 100
    wb.calculation.fullCalcOnLoad = True

    # единый шрифт для листов без явных настроек
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None and cell.font.name != F:
                    cell.font = f_norm
    return wb, layout


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    data, osv_turn = prepare()
    wb, layout = build(data, osv_turn)
    path = OUT / "PnL_2025.xlsx"
    wb.save(path)
    recalculate(path)
    rev = data[data.block == "Выручка"].amount.sum()
    cost = data[data.block == "Себестоимость"].amount.sum()
    mg = data[data.block == "Управленческие"].amount.sum()
    com = data[data.block == "Коммерческие"].amount.sum()
    print(f"Сохранено: {path}  (строк данных: {len(data)})")
    print(f"Контроль в pandas: выручка {rev:,.0f}; себестоимость {cost:,.0f}; управленческие {mg:,.0f}; "
          f"коммерческие {com:,.0f}; прибыль {rev - cost - mg - com:,.0f}")
    print("Строки P&L:", layout)


if __name__ == "__main__":
    main()

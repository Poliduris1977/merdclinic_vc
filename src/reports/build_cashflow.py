"""
ДДС прямым методом и мост «прибыль до налога -> денежный поток»: data/reports/CashFlow_2025.xlsx

Листы:
  Описание     что считает книга, допущения
  ДДС          прямой метод по статьям, помесячно; фильтр «клиника»; остаток на конец сверен с ОСВ сч. 51
  Мост         прибыль до налога -> чистое изменение денег (косвенный метод из ОСВ), помесячно; фильтр общий с ДДС
  Данные ДДС   платежи 1С, агрегат «месяц x клиника x статья» (значения)
  Данные ОСВ   сальдо «месяц x клиника x счёт» (значения)

Месяц в данных хранится числом 1-12: SUMIFS по текстовому «2025-01» Excel может принять за дату.
Счёт хранится как «сч. 51»: SUMIFS по «01» Excel и LibreOffice читают как число 1.
Источник: data/corrected (после исправлений сверки).
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from build_factor import (BLACK, BLUE, BOLD, F, GREEN, GREEN_LINK, INPUT_FILL, NOTE, NUM, RED, SEC_FILL,  # noqa: E402
                          TITLE, TOP, head, sign_format)
from common.recalc import recalculate  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "reports"
MONTHS = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]
CLINICS = ["C01", "C02", "C03"]
TOL = 1.0   # допуск сверки, ₽

IN_ARTICLES = ["Поступления от пациентов", "Поступления от страховых компаний"]
OUT_ARTICLES = ["Выплата заработной платы", "НДФЛ", "Страховые взносы", "Материалы и лаборатории", "Аренда помещений",
                "Коммунальные услуги", "Маркетинг и реклама", "Комиссия эквайринга"]
INV_ARTICLES = ["Приобретение оборудования"]

# (подпись, счёт, примечание); знак всегда «минус изменение сальдо Дт-Кт»
BRIDGE_OP = [
    ("Амортизация (начислена, денег не требует)", "02", "Растёт кредитовое сальдо сч. 02: расход есть, оттока денег нет"),
    ("Изменение запасов материалов", "10", "Рост запасов связывает деньги"),
    ("Изменение дебиторки ДМС", "62.02", "Страховые платят через 30-45 дней: рост долга страховых связывает деньги"),
    ("Изменение расчётов с пациентами", "62.01", "Самооплата в день визита: сальдо около нуля"),
    ("Изменение долга поставщикам", "60", "Рост долга: расход учтён, деньги ещё не уплачены"),
    ("Изменение долга по зарплате", "70", "Зарплата выплачивается со сдвигом"),
    ("Изменение долга по НДФЛ", "68", "Налог уплачивается в следующем месяце"),
    ("Изменение долга по страховым взносам", "69", "Взносы уплачиваются в следующем месяце"),
]


def load():
    cf = pd.read_csv(CORR / "cashflow_1c.csv")
    cf["m"] = cf.date.str[5:7].astype(int)
    dds = cf.groupby(["m", "clinic_id", "activity", "article"], as_index=False).signed_amount.sum()
    dds["signed_amount"] = dds.signed_amount.round(2)
    o = pd.read_csv(CORR / "osv_1c.csv", dtype={"account": str})
    o["m"] = o.period.str[5:7].astype(int)
    o["open_net"] = (o.opening_debit - o.opening_credit).round(2)
    o["close_net"] = (o.closing_debit - o.closing_credit).round(2)
    o["acct"] = "сч. " + o.account
    osv = o.groupby(["m", "clinic_id", "acct"], as_index=False)[["open_net", "close_net"]].sum()
    return dds, osv


def build(dds, osv):
    wb = Workbook()
    nd, no = len(dds), len(osv)
    DD = lambda c: f"'Данные ДДС'!${c}$2:${c}${nd + 1}"  # noqa: E731
    OO = lambda c: f"'Данные ОСВ'!${c}$2:${c}${no + 1}"  # noqa: E731
    crit = "IF($B$2=\"Все\",\"*\",$B$2)"
    M0, MN = 2, 13                      # столбцы B..M месяцы, N итого
    TOT = 14

    # ---------------------------------------------------------- данные
    w1 = wb.active
    w1.title = "Данные ДДС"
    for c, h in enumerate(["Месяц", "Клиника", "Активность", "Статья", "Сумма со знаком, ₽"], 1):
        w1.cell(1, c, h)
    head(w1, 1, 1, 5)
    for i, x in enumerate(dds.itertuples(), 2):
        for c, v in enumerate([int(x.m), x.clinic_id, x.activity, x.article, float(x.signed_amount)], 1):
            w1.cell(i, c, v).font = BLUE
        w1.cell(i, 5).number_format = "#,##0.00"
    for c, w in zip("ABCDE", (8, 9, 12, 36, 18)):
        w1.column_dimensions[c].width = w
    w1.freeze_panes = "A2"

    w2 = wb.create_sheet("Данные ОСВ")
    for c, h in enumerate(["Месяц", "Клиника", "Счёт", "Сальдо на начало (Дт−Кт), ₽", "Сальдо на конец (Дт−Кт), ₽"], 1):
        w2.cell(1, c, h)
    head(w2, 1, 1, 5)
    for i, x in enumerate(osv.itertuples(), 2):
        for c, v in enumerate([int(x.m), x.clinic_id, x.acct, float(x.open_net), float(x.close_net)], 1):
            w2.cell(i, c, v).font = BLUE
        w2.cell(i, 4).number_format = w2.cell(i, 5).number_format = "#,##0.00"
    for c, w in zip("ABCDE", (8, 9, 12, 26, 26)):
        w2.column_dimensions[c].width = w
    w2.freeze_panes = "A2"

    # общая шапка листа с месяцами
    def header(ws, title, sub):
        ws["A1"], ws["A1"].font = title, TITLE
        ws["A2"], ws["A2"].font = "Клиника (фильтр):", BOLD
        ws["A3"], ws["A3"].font = sub, NOTE
        for c, h in enumerate(["Статья"] + MONTHS + ["Итого 2025"], 1):
            ws.cell(4, c, h)
        head(ws, 4, 1, TOT)
        ws.cell(5, 1, "№ месяца (служебная строка для формул)").font = NOTE
        for k in range(12):
            cell = ws.cell(5, M0 + k, k + 1)
            cell.font = NOTE
            cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions["A"].width = 52
        for c in range(M0, TOT + 1):
            ws.column_dimensions[col(c)].width = 11 if c < TOT else 13
        ws.freeze_panes = "B6"

    def section(ws, r, text):
        for c in range(1, TOT + 1):
            ws.cell(r, c).fill = SEC_FILL
        ws.cell(r, 1, text).font = BOLD

    def row_fmt(ws, r, bold=False, top=False):
        for c in range(1, TOT + 1):
            cell = ws.cell(r, c)
            if c > 1:
                cell.number_format = NUM
            if bold:
                cell.font = BOLD
            if top:
                cell.border = TOP

    # ---------------------------------------------------------- ДДС
    ws = wb.create_sheet("ДДС", 0)
    header(ws, "Отчёт о движении денежных средств, прямой метод, 2025",
           "Тыс. ₽. Поступления положительные, выплаты отрицательные. Фильтр: «Все» или C01 / C02 / C03.")
    ws["B2"], ws["B2"].fill, ws["B2"].font = "Все", INPUT_FILL, BLUE
    dv = DataValidation(type="list", formula1='"Все,C01,C02,C03"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B2")

    def art_row(r, art):
        ws.cell(r, 1, art).font = BLACK
        for c in range(M0, MN + 1):
            L = col(c)
            ws.cell(r, c, f"=SUMIFS({DD('E')},{DD('A')},{L}$5,{DD('B')},{crit},{DD('D')},$A{r})").font = GREEN_LINK
        ws.cell(r, TOT, f"=SUM(B{r}:M{r})").font = BLACK
        row_fmt(ws, r)

    r = 6
    section(ws, r, "Операционная деятельность")
    r += 1
    op_first = r
    for a in IN_ARTICLES + OUT_ARTICLES:
        art_row(r, a)
        r += 1
    op_last = r - 1
    op_tot = r
    ws.cell(r, 1, "Чистый операционный денежный поток")
    for c in range(M0, TOT + 1):
        ws.cell(r, c, f"=SUM({col(c)}{op_first}:{col(c)}{op_last})")
    row_fmt(ws, r, bold=True, top=True)
    r += 2
    section(ws, r, "Инвестиционная деятельность")
    r += 1
    inv_first = r
    for a in INV_ARTICLES:
        art_row(r, a)
        r += 1
    inv_tot = r
    ws.cell(r, 1, "Чистый инвестиционный денежный поток")
    for c in range(M0, TOT + 1):
        ws.cell(r, c, f"=SUM({col(c)}{inv_first}:{col(c)}{r - 1})")
    row_fmt(ws, r, bold=True, top=True)
    r += 2
    net = r
    ws.cell(r, 1, "Чистое изменение денежных средств")
    for c in range(M0, TOT + 1):
        ws.cell(r, c, f"={col(c)}{op_tot}+{col(c)}{inv_tot}")
    row_fmt(ws, r, bold=True, top=True)
    sign_format(ws, f"B{net}:N{net}")
    r += 1
    opn = r
    ws.cell(r, 1, "Остаток на начало периода")
    ws.cell(r, M0, f"=SUMIFS({OO('D')},{OO('A')},B$5,{OO('B')},{crit},{OO('C')},\"сч. 51\")").font = GREEN_LINK
    for c in range(M0 + 1, MN + 1):
        ws.cell(r, c, f"={col(c - 1)}{r + 1}").font = BLACK
    ws.cell(r, TOT, f"=B{r}").font = BLACK
    row_fmt(ws, r)
    r += 1
    cl = r
    ws.cell(r, 1, "Остаток на конец периода").font = BOLD
    for c in range(M0, MN + 1):
        ws.cell(r, c, f"={col(c)}{opn}+{col(c)}{net}")
    ws.cell(r, TOT, f"=M{r}")
    row_fmt(ws, r, bold=True, top=True)
    r += 2
    section(ws, r, "Сверка с ОСВ (счёт 51 «Расчётные счета»)")
    r += 1
    osv_r = r
    ws.cell(r, 1, "Остаток на конец по ОСВ, сч. 51").font = BLACK
    for c in range(M0, MN + 1):
        ws.cell(r, c, f"=SUMIFS({OO('E')},{OO('A')},{col(c)}$5,{OO('B')},{crit},{OO('C')},\"сч. 51\")").font = GREEN_LINK
    ws.cell(r, TOT, f"=M{r}").font = BLACK
    row_fmt(ws, r)
    r += 1
    diff = r
    ws.cell(r, 1, "Расхождение (ДДС − ОСВ), ₽").font = BLACK
    for c in range(M0, TOT + 1):
        ws.cell(r, c, f"=ROUND({col(c)}{cl}-{col(c)}{osv_r},2)").font = BLACK
        ws.cell(r, c).number_format = "#,##0.00;-#,##0.00;0.00"
    r += 1
    ws.cell(r, 1, "Статус").font = BOLD
    ws.cell(r, 2, f'=IF(SUMPRODUCT(--(ABS(B{diff}:M{diff})>{TOL}))=0,"СХОДИТСЯ","РАСХОЖДЕНИЕ")').font = BOLD
    ws.cell(r + 1, 1, f"Допуск {TOL:.0f} ₽ на месяц. Остаток на начало января берётся из ОСВ, дальше накапливается по ДДС.").font = NOTE
    status_dds = r

    ch = BarChart()
    ch.type, ch.title, ch.height, ch.width = "col", "Чистое изменение денежных средств по месяцам, ₽", 7.5, 22
    ch.add_data(Reference(ws, min_col=M0, max_col=MN, min_row=net), from_rows=True, titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=M0, max_col=MN, min_row=4))
    ch.legend = None
    ws.add_chart(ch, f"A{r + 4}")

    # ---------------------------------------------------------- Мост
    wm = wb.create_sheet("Мост", 1)
    header(wm, "Мост: прибыль до налога → денежный поток, 2025",
           "Тыс. ₽. Расчёт из ОСВ: изменение сальдо балансовых счетов с обратным знаком. Фильтр клиники задаётся на листе «ДДС».")
    wm["B2"], wm["B2"].font = "=ДДС!B2", BLACK
    crit_m = "IF(ДДС!$B$2=\"Все\",\"*\",ДДС!$B$2)"

    def delta(L, acct):   # -(сальдо на конец - сальдо на начало) по счёту за месяц
        return (f"=-(SUMIFS({OO('E')},{OO('A')},{L}$5,{OO('B')},{crit_m},{OO('C')},\"сч. {acct}\")"
                f"-SUMIFS({OO('D')},{OO('A')},{L}$5,{OO('B')},{crit_m},{OO('C')},\"сч. {acct}\"))")

    def brow(r, label, acct, note=None, bold=False):
        wm.cell(r, 1, label).font = BOLD if bold else BLACK
        for c in range(M0, MN + 1):
            wm.cell(r, c, delta(col(c), acct)).font = GREEN_LINK
        wm.cell(r, TOT, f"=SUM(B{r}:M{r})").font = BLACK
        row_fmt(wm, r)
        if note:
            wm.cell(r, TOT + 1, note).font = NOTE

    wm.column_dimensions[col(TOT + 1)].width = 70
    wm.cell(4, TOT + 1, "Пояснение")
    head(wm, 4, TOT + 1, TOT + 1)
    r = 6
    section(wm, r, "Операционная деятельность")
    r += 1
    prof = r
    brow(r, "Прибыль до налога (сч. 99)", "99", "Прибыль месяца: закрытие месяца относит результат на сч. 99")
    wm.cell(r, 1).font = BOLD
    r += 1
    first_adj = r
    for label, acct, note in BRIDGE_OP:
        brow(r, label, acct, note)
        r += 1
    last_adj = r - 1
    opb = r
    wm.cell(r, 1, "Операционный денежный поток (по мосту)")
    for c in range(M0, TOT + 1):
        wm.cell(r, c, f"=SUM({col(c)}{prof}:{col(c)}{last_adj})")
    row_fmt(wm, r, bold=True, top=True)
    r += 2
    section(wm, r, "Инвестиционная деятельность")
    r += 1
    capex = r
    brow(r, "Капитальные вложения (рост сч. 01)", "01", "Покупка оборудования в прибыль не попадает, кроме амортизации")
    r += 1
    netb = r
    wm.cell(r, 1, "Чистое изменение денежных средств (по мосту)")
    for c in range(M0, TOT + 1):
        wm.cell(r, c, f"={col(c)}{opb}+{col(c)}{capex}")
    row_fmt(wm, r, bold=True, top=True)
    sign_format(wm, f"B{prof}:N{netb}")
    r += 2
    section(wm, r, "Сверка")
    r += 1
    d51 = r
    wm.cell(r, 1, "Изменение остатка на сч. 51 по ОСВ").font = BLACK
    for c in range(M0, MN + 1):
        L = col(c)
        wm.cell(r, c, f"=SUMIFS({OO('E')},{OO('A')},{L}$5,{OO('B')},{crit_m},{OO('C')},\"сч. 51\")"
                      f"-SUMIFS({OO('D')},{OO('A')},{L}$5,{OO('B')},{crit_m},{OO('C')},\"сч. 51\")").font = GREEN_LINK
    wm.cell(r, TOT, f"=SUM(B{r}:M{r})").font = BLACK
    row_fmt(wm, r)
    r += 1
    d_br = r
    wm.cell(r, 1, "Расхождение: мост − ОСВ, ₽").font = BLACK
    for c in range(M0, TOT + 1):
        wm.cell(r, c, f"=ROUND({col(c)}{netb}-{col(c)}{d51},2)").font = BLACK
        wm.cell(r, c).number_format = "#,##0.00;-#,##0.00;0.00"
    r += 1
    d_cf = r
    wm.cell(r, 1, "Расхождение: мост − ДДС (прямой метод), ₽").font = BLACK
    for c in range(M0, TOT + 1):
        wm.cell(r, c, f"=ROUND({col(c)}{netb}-ДДС!{col(c)}{net},2)").font = BLACK
        wm.cell(r, c).number_format = "#,##0.00;-#,##0.00;0.00"
    r += 1
    wm.cell(r, 1, "Статус").font = BOLD
    wm.cell(r, 2, f'=IF(SUMPRODUCT(--(ABS(B{d_br}:M{d_cf})>{TOL}))=0,"СХОДИТСЯ","РАСХОЖДЕНИЕ")').font = BOLD
    status_br = r
    r += 2
    wm.cell(r, 1, "Прибыль до налога и денежный поток отличаются из-за трёх вещей: неденежных расходов (амортизация), "
                  "разницы во времени между начислением и оплатой (дебиторка ДМС, долги по зарплате, налогам, поставщикам) "
                  "и капитальных вложений, которых нет в прибыли.").font = NOTE

    # ---------------------------------------------------------- Описание
    wo = wb.create_sheet("Описание", 0)
    wo["A1"], wo["A1"].font = "Описание", TITLE
    lines = [
        ("Что это: движение денег сети клиник за 2025 год (синтетические данные, демо). Источник: данные после исправлений сверки (data/corrected).", BLACK),
        ("", BLACK),
        ("Лист «ДДС»: прямой метод по статьям платежей 1С, помесячно. Фильтр клиники в жёлтой ячейке B2. "
         "Остаток на конец сверяется с ОСВ счёта 51 (допуск 1 ₽ на месяц).", BLACK),
        ("Лист «Мост»: от прибыли до налога к чистому изменению денег. Считается из сальдо ОСВ независимо от ДДС; "
         "контроль: мост равен изменению счёта 51 и равен ДДС.", BLACK),
        ("", BLACK),
        ("Допущения и ограничения:", BOLD),
        ("1. Нет финансовой деятельности (кредитов, дивидендов) и налога на прибыль: поток заканчивается инвестициями, прибыль считается до налога.", BLACK),
        ("2. Нет плана по деньгам: бюджет составлен только по P&L, поэтому план-факт ДДС не строится.", BLACK),
        ("3. Входящие остатки (на 01.01.2025) заданы генератором, остаток на конец года 216 млн ₽ зависит от них.", BLACK),
        ("4. Дебиторка ДМС в ОСВ одним сальдо без разбивки по страховым и срокам: «возраст» долга не определяется.", BLACK),
        ("5. Деньги на счёте клиник считаются отдельно для каждой клиники; движений между клиниками нет.", BLACK),
        ("6. Страховые взносы и НДФЛ платятся в следующем месяце, зарплата выплачивается со сдвигом (так задано в генераторе 1С).", BLACK),
        ("", BLACK),
        ("Цвета: синий шрифт: входные числа; зелёный: ссылка на листы данных; чёрный: расчёты; жёлтая заливка: фильтр.", NOTE),
    ]
    for i, (t, f) in enumerate(lines, 3):
        c = wo.cell(i, 1, t)
        c.font = f
        c.alignment = Alignment(wrap_text=True, vertical="top")
    wo.column_dimensions["A"].width = 130
    return wb


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    dds, osv = load()
    wb = build(dds, osv)
    path = OUT / "CashFlow_2025.xlsx"
    wb.save(path)
    recalculate(path)
    print(f"Сохранено: {path}  (строк ДДС: {len(dds)}, ОСВ: {len(osv)})")
    by = dds.groupby("activity").signed_amount.sum()
    print(f"Контроль pandas: операционный {by['operating']:,.0f}; инвестиционный {by['investing']:,.0f}; чистое изменение {by.sum():,.0f} ₽")


if __name__ == "__main__":
    main()

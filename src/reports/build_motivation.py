"""
Калькулятор надбавок за KPI врачей: data/reports/MotivationCalc_2025.xlsx

Листы:
  Описание     что считает книга, допущения
  Калькулятор  параметры (жёлтые ячейки), сводка по сети и по направлениям
  Врачи        43 врача: входные числа 2025 года (синий), KPI и надбавка (формулы)

KPI 1: доля пациентов, вернувшихся за 90 дней, минус средняя по направлению (п.п.).
KPI 2: средний чек врача к среднему чеку специальности (чек зависит от специальности, не от направления).
Надбавка добавляется к доле врача от выручки (п.п.), взносы начисляются сверху.
Расчёт показывает СТОИМОСТЬ схемы на данных 2025 года, а не изменение поведения врачей.
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from build_doctors import NODOC, prepare  # noqa: E402
from build_factor import (BLACK, BLUE, BOLD, DIRS, GREEN, INPUT_FILL, NOTE, NUM, PCT, RED,  # noqa: E402
                          SEC_FILL, TITLE, TOP, head, sign_format)
from common.recalc import recalculate  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "data" / "reports"
PPF = '0.0%;-0.0%;0.0%'
MLN = '#,##0.00,,;(#,##0.00,,);-'

# параметры по умолчанию (вход, меняется в книге)
DEFAULTS = dict(thr1=0.02, bonus1=0.01, thr2=0.05, bonus2=0.01, min_cohort=100, rate=0.22, stack="Да")


def doctor_table():
    data, doctors, ret, _, _ = prepare()
    d = data[data.doctor_id != NODOC].groupby("doctor_id", as_index=False).agg(
        visits=("visits", "sum"), revenue=("revenue", "sum"), norm_cost=("norm_cost", "sum"),
        pay_total=("pay_total", "sum"))
    d = d.merge(doctors[["doctor_id", "doctor_name", "direction_code", "direction", "specialty", "clinic_id", "grade"]], on="doctor_id")
    d = d.merge(ret, on="doctor_id", how="left").fillna({"cohort_n": 0, "returned": 0})
    d["spec_key"] = d.direction + ": " + d.specialty
    d = d.sort_values(["direction_code", "spec_key", "doctor_id"]).reset_index(drop=True)
    return d


def control(d, p):
    """Независимый расчёт в pandas для проверки формул Excel."""
    d = d.copy()
    d["ret_n"] = d.cohort_n * d.returned
    avg = d.groupby("direction_code").ret_n.sum() / d.groupby("direction_code").cohort_n.sum()
    d["k1"] = (d.cohort_n >= p["min_cohort"]) & (d.returned - d.direction_code.map(avg) >= p["thr1"])
    chk = d.revenue / d.visits
    sp = d.groupby("spec_key").revenue.sum() / d.groupby("spec_key").visits.sum()
    d["k2"] = chk / d.spec_key.map(sp) >= 1 + p["thr2"]
    a, b = d.k1 * p["bonus1"], d.k2 * p["bonus2"]
    d["bonus_pp"] = a + b if p["stack"] == "Да" else a.combine(b, max)
    d["cost"] = d.bonus_pp * d.revenue * (1 + p["rate"])
    return d


def build(d):
    wb = Workbook()
    n = len(d)
    r1, r2 = 6, 5 + n          # строки врачей на листе «Врачи»
    V = lambda c: f"'Врачи'!${c}${r1}:${c}${r2}"  # noqa: E731

    # ------------------------------------------------------------ Врачи
    wv = wb.active
    wv.title = "Врачи"
    wv["A1"], wv["A1"].font = "Врачи: KPI и надбавка, 2025", TITLE
    wv["A2"] = ("Синий: входные числа 2025 года (визиты, выручка, оплата из 1С, когорта повторных визитов). Чёрный: формулы, "
                "зависят от параметров на листе «Калькулятор». Сравнение чека: внутри специальности; повторных визитов: внутри направления.")
    wv["A3"] = ("Малая выборка: у специальности 1-6 врачей, у специальности с одним врачом чек равен среднему и KPI 2 не срабатывает. "
                "Без «Без врача» (2 визита, 120 700 ₽).")
    wv["A2"].font = wv["A3"].font = NOTE
    heads = ["Врач", "ФИО", "Код напр.", "Направление: специальность", "Клиника", "Грейд", "Визитов", "Выручка, тыс. ₽",
             "Материалы и лаб. (норматив), тыс. ₽", "Оплата врача с взносами (факт), тыс. ₽", "Когорта пациентов",
             "Вернулись за 90 дней", "Вернулись, чел.", "Среднее по направлению", "Отклонение, п.п.", "KPI 1 выполнен",
             "Чек, ₽", "Средний чек специальности, ₽", "Чек к среднему по специальности", "KPI 2 выполнен",
             "Надбавка к доле, п.п.", "Стоимость надбавки с взносами, тыс. ₽", "Оплата по новой схеме, тыс. ₽",
             "Доля оплаты в выручке: сейчас", "Доля оплаты в выручке: новая", "Маржа после ФОТ: сейчас, тыс. ₽",
             "Маржа после ФОТ: новая, тыс. ₽"]
    for c, h in enumerate(heads, 1):
        wv.cell(5, c, h)
    head(wv, 5, 1, len(heads))
    wv.row_dimensions[5].height = 54
    for i, x in d.iterrows():
        r = r1 + i
        vals = [x.doctor_id, x.doctor_name, x.direction_code, x.spec_key, x.clinic_id, x.grade, int(x.visits),
                float(x.revenue), float(x.norm_cost), float(x.pay_total), int(x.cohort_n), float(x.returned)]
        for c, v in enumerate(vals, 1):
            cell = wv.cell(r, c, v)
            cell.font = BLUE if c >= 7 else BLACK
        f = {
            13: f"=K{r}*L{r}",
            14: f"=IFERROR(SUMIFS($M${r1}:$M${r2},$C${r1}:$C${r2},C{r})/SUMIFS($K${r1}:$K${r2},$C${r1}:$C${r2},C{r}),0)",
            15: f"=(L{r}-N{r})*100",
            16: f"=IF(AND(K{r}>=Калькулятор!$B$9,L{r}-N{r}>=Калькулятор!$B$5),1,0)",
            17: f"=IFERROR(H{r}/G{r},0)",
            18: f"=IFERROR(SUMIFS($H${r1}:$H${r2},$D${r1}:$D${r2},D{r})/SUMIFS($G${r1}:$G${r2},$D${r1}:$D${r2},D{r}),0)",
            19: f"=IFERROR(Q{r}/R{r},0)",
            20: f"=IF(S{r}>=1+Калькулятор!$B$7,1,0)",
            21: f'=IF(Калькулятор!$B$11="Да",P{r}*Калькулятор!$B$6+T{r}*Калькулятор!$B$8,MAX(P{r}*Калькулятор!$B$6,T{r}*Калькулятор!$B$8))*100',
            22: f"=U{r}/100*H{r}*(1+Калькулятор!$B$10)",
            23: f"=J{r}+V{r}",
            24: f"=IFERROR(J{r}/H{r},0)",
            25: f"=IFERROR(W{r}/H{r},0)",
            26: f"=H{r}-I{r}-J{r}",
            27: f"=Z{r}-V{r}",
        }
        for c, fx in f.items():
            wv.cell(r, c, fx).font = BLACK
        for c in (8, 9, 10, 22, 23, 26, 27):
            wv.cell(r, c).number_format = NUM
        for c in (12, 14, 24, 25):
            wv.cell(r, c).number_format = PCT
        wv.cell(r, 15).number_format = '0.0;-0.0;0.0'
        wv.cell(r, 21).number_format = '0.0;-0.0;-'
        wv.cell(r, 17).number_format = wv.cell(r, 18).number_format = "#,##0"
        wv.cell(r, 19).number_format = "0.00"
        wv.cell(r, 7).number_format = wv.cell(r, 11).number_format = wv.cell(r, 13).number_format = "#,##0"
    tr = r2 + 1
    wv.cell(tr, 1, "Итого").font = BOLD
    for c in (7, 8, 9, 10, 11, 13, 22, 23, 26, 27):
        cell = wv.cell(tr, c, f"=SUM({col(c)}{r1}:{col(c)}{r2})")
        cell.font = BOLD
        cell.number_format = NUM if c in (8, 9, 10, 22, 23, 26, 27) else "#,##0"
    wv.cell(tr, 16, f"=SUM(P{r1}:P{r2})").font = BOLD
    wv.cell(tr, 20, f"=SUM(T{r1}:T{r2})").font = BOLD
    wv.cell(tr, 24, f"=J{tr}/H{tr}").number_format = PCT
    wv.cell(tr, 25, f"=W{tr}/H{tr}").number_format = PCT
    for c in range(1, len(heads) + 1):
        wv.cell(tr, c).border = TOP
    widths = [7, 15, 8, 36, 8, 8, 9, 12, 14, 15, 11, 11, 10, 12, 11, 9, 9, 12, 12, 9, 11, 14, 14, 12, 12, 14, 14]
    for c, w in enumerate(widths, 1):
        wv.column_dimensions[col(c)].width = w
    wv.freeze_panes = "C6"
    wv.auto_filter.ref = f"A5:{col(len(heads))}{r2}"
    sign_format(wv, f"V{r1}:V{r2}")

    # ------------------------------------------------------------ Калькулятор
    ws = wb.create_sheet("Калькулятор", 0)
    ws["A1"], ws["A1"].font = "Калькулятор надбавок за KPI врачей, 2025", TITLE
    ws["A2"] = ("Меняйте жёлтые ячейки: сводка и лист «Врачи» пересчитаются. Это стоимость схемы на фактических данных 2025 года; "
                "поведение врачей в расчёте не меняется.")
    ws["A2"].font = NOTE
    ws.cell(4, 1, "Параметры").font = BOLD
    for c in (1, 2, 3):
        ws.cell(4, c).fill = SEC_FILL
    params = [
        (5, "KPI 1: порог, повторные визиты выше среднего по направлению, п.п.", DEFAULTS["thr1"], PPF,
         "Вернувшиеся за 90 дней в то же направление, среди пациентов, впервые пришедших к врачу до 02.10.2025"),
        (6, "KPI 1: надбавка к доле врача, п.п.", DEFAULTS["bonus1"], PPF, "Например, 1,0% значит 36% превращаются в 37%"),
        (7, "KPI 2: порог, чек выше среднего по специальности, %", DEFAULTS["thr2"], PPF, "Чек врача к среднему чеку его специальности"),
        (8, "KPI 2: надбавка к доле врача, п.п.", DEFAULTS["bonus2"], PPF, ""),
        (9, "KPI 1: минимальная когорта пациентов у врача", DEFAULTS["min_cohort"], "#,##0", "Меньшая когорта: KPI 1 не считается (шум)"),
        (10, "Ставка взносов на надбавку", DEFAULTS["rate"], PPF, "Допущение демо: плоские 22%, как в остальных отчётах"),
        (11, "Надбавки суммируются (Да) или берётся большая (Нет)", DEFAULTS["stack"], "@", "Если врач выполнил оба KPI"),
    ]
    for r, label, v, fmt, note in params:
        ws.cell(r, 1, label).font = BLACK
        cell = ws.cell(r, 2, v)
        cell.font, cell.fill, cell.number_format = BLUE, INPUT_FILL, fmt
        ws.cell(r, 3, note).font = NOTE
    dv = DataValidation(type="list", formula1='"Да,Нет"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B11")

    ws.cell(13, 1, "Сводка по сети (43 врача)").font = BOLD
    for c in (1, 2, 3):
        ws.cell(13, c).fill = SEC_FILL
    T = tr
    rows = [
        (14, "Выручка врачей, млн ₽", f"='Врачи'!H{T}", MLN, "Без двух визитов без врача"),
        (15, "Оплата врачей с взносами: сейчас, млн ₽", f"='Врачи'!J{T}", MLN, "Факт по 1С"),
        (16, "Стоимость надбавок с взносами, млн ₽", f"='Врачи'!V{T}", MLN, ""),
        (17, "Оплата по новой схеме, млн ₽", f"='Врачи'!W{T}", MLN, ""),
        (18, "Рост оплаты врачей, %", "=IFERROR(B16/B15,0)", PCT, ""),
        (19, "Доля оплаты в выручке: сейчас", f"='Врачи'!X{T}", PCT, ""),
        (20, "Доля оплаты в выручке: новая", f"='Врачи'!Y{T}", PCT, ""),
        (21, "Маржа после ФОТ: сейчас, % выручки", f"='Врачи'!Z{T}/'Врачи'!H{T}", PCT, "Выручка − материалы и лаб. (норматив) − оплата с взносами"),
        (22, "Маржа после ФОТ: новая, % выручки", f"='Врачи'!AA{T}/'Врачи'!H{T}", PCT, ""),
        (23, "Врачей выполнили KPI 1 / KPI 2 / хотя бы один",
         f"=SUM({V('P')})&\" / \"&SUM({V('T')})&\" / \"&COUNTIF({V('U')},\">0\")", "@", "Из 43"),
    ]
    for r, label, fx, fmt, note in rows:
        ws.cell(r, 1, label).font = BLACK
        cell = ws.cell(r, 2, fx)
        cell.font, cell.number_format = BLACK, fmt
        cell.alignment = Alignment(horizontal="right")
        ws.cell(r, 3, note).font = NOTE
    ws.cell(16, 2).font = ws.cell(18, 2).font = BOLD

    hr = 26
    ws.cell(hr - 1, 1, "По направлениям").font = BOLD
    for c in range(1, 9):
        ws.cell(hr - 1, c).fill = SEC_FILL
    for c, h in enumerate(["Направление", "Врачей", "Выполнили KPI 1", "Выполнили KPI 2", "Выручка, млн ₽",
                           "Стоимость надбавок, млн ₽", "Надбавки, % оплаты врачей", "Маржа после ФОТ: изменение, п.п."], 1):
        ws.cell(hr, c, h)
    head(ws, hr, 1, 8)
    ws.row_dimensions[hr].height = 42
    for i, (code, name) in enumerate(DIRS.items()):
        r = hr + 1 + i
        ws.cell(r, 1, name).font = BLACK
        sumif = lambda c: f"SUMIFS({V(c)},{V('C')},\"{code}\")"  # noqa: E731
        fx = {
            2: f"=COUNTIF({V('C')},\"{code}\")",
            3: f"={sumif('P')}",
            4: f"={sumif('T')}",
            5: f"={sumif('H')}",
            6: f"={sumif('V')}",
            7: f"=IFERROR({sumif('V')}/{sumif('J')},0)",
            8: f"=IFERROR(({sumif('AA')}-{sumif('Z')})/{sumif('H')},0)*100",
        }
        for c, v in fx.items():
            ws.cell(r, c, v).font = BLACK
        ws.cell(r, 5).number_format = ws.cell(r, 6).number_format = MLN
        ws.cell(r, 7).number_format = PCT
        ws.cell(r, 8).number_format = '0.00;-0.00;-'
    last = hr + len(DIRS)
    tr2 = last + 1
    ws.cell(tr2, 1, "Итого").font = BOLD
    for c in (2, 3, 4, 5, 6):
        cell = ws.cell(tr2, c, f"=SUM({col(c)}{hr + 1}:{col(c)}{last})")
        cell.font = BOLD
        cell.number_format = MLN if c in (5, 6) else "0"
    ws.cell(tr2, 7, "=B18").number_format = PCT
    ws.cell(tr2, 8, "=(B22-B21)*100").number_format = '0.00;-0.00;-'
    for c in range(1, 9):
        ws.cell(tr2, c).border = TOP
    ws.column_dimensions["A"].width = 58
    for c, w in zip("BCDEFGH", (16, 16, 16, 14, 16, 16, 18)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A4"

    # ------------------------------------------------------------ Описание
    wo = wb.create_sheet("Описание", 0)
    wo["A1"], wo["A1"].font = "Описание", TITLE
    lines = [
        ("Что это: калькулятор стоимости надбавок за KPI врачей на данных 2025 года (синтетическая сеть клиник, демо).", BLACK),
        ("Как пользоваться: лист «Калькулятор», жёлтые ячейки B5:B11. Сводка и таблица «Врачи» пересчитываются сразу.", BLACK),
        ("", BLACK),
        ("KPI 1. Повторные визиты: доля пациентов, вернувшихся в то же направление за 90 дней, минус средняя по направлению (взвешена по когорте). "
         "Сравнение внутри направления, потому что между направлениями разница в разы (семейная около 50%, лазер и пластика около 5%).", BLACK),
        ("KPI 2. Чек: средний чек врача к среднему чеку его специальности (ортопед и терапевт в одной стоматологии несопоставимы).", BLACK),
        ("Надбавка: прибавляется к доле врача от выручки (п.п.), взносы начисляются сверху. Расчёт идёт на фактической выручке 2025 года.", BLACK),
        ("", BLACK),
        ("Допущения и ограничения:", BOLD),
        ("1. Это стоимость схемы, а не прогноз. Реакция врачей на надбавки (больше возвратов, другой чек) в расчёт не входит.", BLACK),
        ("2. В синтетических данных врач почти не влияет на возврат пациентов: внутри направления разброс повторных визитов 2-3 п.п. "
         "Результат по KPI 1 в основном шум; в реальных данных разброс будет шире.", BLACK),
        ("3. Чек зависит от состава услуг, а не только от качества врача; внутри специальности он тоже может отражать загрузку и клинику.", BLACK),
        ("4. Малая выборка: 6-12 врачей на направление, у некоторых специальностей один врач. Выводы по конкретному врачу делать нельзя.", BLACK),
        ("5. Взносы на надбавку: плоские 22% (допущение демо, как в остальных отчётах).", BLACK),
        ("6. База: выручка врача. Доли по грейдам (28/32/36/40%) и фиксированная часть не меняются.", BLACK),
        ("7. «Маржа после ФОТ» использует материалы и лаборатории по нормативу услуг, а не по факту 1С.", BLACK),
        ("8. Не входят 2 визита без врача (120 700 ₽).", BLACK),
        ("", BLACK),
        ("Цвета: синий шрифт: входные числа; чёрный: формулы; жёлтая заливка: параметры.", NOTE),
    ]
    for i, (t, f) in enumerate(lines, 3):
        c = wo.cell(i, 1, t)
        c.font = f
        c.alignment = Alignment(wrap_text=True, vertical="top")
    wo.column_dimensions["A"].width = 130
    return wb


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    d = doctor_table()
    wb = build(d)
    path = OUT / "MotivationCalc_2025.xlsx"
    wb.save(path)
    recalculate(path)
    c = control(d, DEFAULTS)
    print(f"Сохранено: {path}  (врачей: {len(d)})")
    print(f"Контроль pandas: выручка {d.revenue.sum():,.0f}; оплата с взносами {d.pay_total.sum():,.0f}; "
          f"KPI1 {int(c.k1.sum())}, KPI2 {int(c.k2.sum())}, любой {int((c.bonus_pp > 0).sum())}; "
          f"стоимость надбавок {c.cost.sum():,.0f} ₽ (+{c.cost.sum() / d.pay_total.sum():.2%} к оплате)")


if __name__ == "__main__":
    main()

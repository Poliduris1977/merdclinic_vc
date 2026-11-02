"""
Общие функции работы с бухгалтерским слоем 1С: закрытие месяца, ОСВ, ДДС.

Используются и генератором ошибок (inject_errors.py), и модулем исправлений
(apply_corrections.py): после любой правки проводок закрытие месяца, ОСВ и ДДС
пересчитываются здесь, поэтому файлы остаются внутренне согласованными.

ACCOUNTS и CASH_MAP - план счетов и справочник статей ДДС; значения совпадают
с generate_1c.py (там они используются при генерации).
"""
import calendar
from datetime import date

import numpy as np
import pandas as pd

YEAR = 2025

ACCOUNTS = {'01': 'Основные средства',
 '02': 'Амортизация основных средств',
 '10': 'Материалы',
 '20': 'Основное производство',
 '26': 'Общехозяйственные расходы',
 '44': 'Расходы на продажу',
 '51': 'Расчётные счета',
 '60': 'Расчёты с поставщиками и подрядчиками',
 '62.01': 'Расчёты с пациентами',
 '62.02': 'Расчёты со страховыми компаниями (ДМС)',
 '68': 'Расчёты по налогам (НДФЛ)',
 '69': 'Расчёты по страховым взносам',
 '70': 'Расчёты с персоналом по оплате труда',
 '80': 'Капитал (входящие остатки)',
 '90.01': 'Выручка',
 '90.02': 'Себестоимость продаж',
 '90.07': 'Расходы на продажу (закрытие)',
 '90.08': 'Управленческие расходы (закрытие)',
 '90.09': 'Прибыль/убыток от продаж',
 '99': 'Прибыли и убытки'}

CASH_MAP = {'Поступление от пациентов (эквайринг)': ('operating', 'Поступления от пациентов', 'in'),
 'Поступление от страховой': ('operating', 'Поступления от страховых компаний', 'in'),
 'Комиссия эквайринга': ('operating', 'Комиссия эквайринга', 'out'),
 'Выплата зарплаты': ('operating', 'Выплата заработной платы', 'out'),
 'Уплата НДФЛ': ('operating', 'НДФЛ', 'out'),
 'Уплата страховых взносов': ('operating', 'Страховые взносы', 'out'),
 'Оплата аренды': ('operating', 'Аренда помещений', 'out'),
 'Оплата коммунальных услуг': ('operating', 'Коммунальные услуги', 'out'),
 'Оплата рекламы': ('operating', 'Маркетинг и реклама', 'out'),
 'Оплата поставщикам': ('operating', 'Материалы и лаборатории', 'out'),
 'Оплата оборудования': ('investing', 'Приобретение оборудования', 'out')}


def month_end(y, m):
    return date(y, m, calendar.monthrange(y, m)[1])


def rebuild_closing(oper: pd.DataFrame) -> pd.DataFrame:
    """Закрытие месяцев (90 -> 99) по операционным проводкам."""
    p = oper.copy()
    p["month"] = pd.to_datetime(p.date).dt.month
    rows = []

    def add(d, cid, dr, cr, amt, direction=""):
        amt = round(float(amt), 2)
        if amt:
            rows.append(dict(date=d.isoformat(), clinic_id=cid, debit=dr, credit=cr, amount=amt,
                             direction_code=direction, counterparty="", doc_type="Закрытие месяца",
                             comment=""))

    c20 = p[p.debit == "20"].groupby(["month", "clinic_id", "direction_code"]).amount.sum()
    c26 = p[p.debit == "26"].groupby(["month", "clinic_id"]).amount.sum()
    c44 = p[p.debit == "44"].groupby(["month", "clinic_id"]).amount.sum()
    rev = p[p.credit == "90.01"].groupby(["month", "clinic_id"]).amount.sum()
    for (m, cid), r_amt in rev.items():
        last = month_end(YEAR, m)
        cost20 = 0.0
        for (mm, cc, dcode), v in c20.items():
            if mm == m and cc == cid:
                add(last, cid, "90.02", "20", v, dcode)
                cost20 += v
        v44, v26 = c44.get((m, cid), 0), c26.get((m, cid), 0)
        add(last, cid, "90.07", "44", v44)
        add(last, cid, "90.08", "26", v26)
        add(last, cid, "90.01", "90.09", r_amt)
        add(last, cid, "90.09", "90.02", cost20)
        add(last, cid, "90.09", "90.07", v44)
        add(last, cid, "90.09", "90.08", v26)
        profit = round(r_amt - cost20 - v44 - v26, 2)
        if profit >= 0:
            add(last, cid, "90.09", "99", profit)
        else:
            add(last, cid, "99", "90.09", -profit)
    return pd.DataFrame(rows)


def rebuild_osv(post: pd.DataFrame, osv_clean: pd.DataFrame) -> pd.DataFrame:
    """ОСВ из проводок и входящих остатков (берутся из чистой ОСВ за январь)."""
    jan = osv_clean[osv_clean.period == f"{YEAR}-01"]
    openings = {(r.clinic_id, r.account): r.opening_debit - r.opening_credit for r in jan.itertuples()}
    p = post.copy()
    p["month"] = pd.to_datetime(p.date).dt.month
    clinics = sorted(p.clinic_id.unique())
    idx = pd.MultiIndex.from_product([range(1, 13), clinics, list(ACCOUNTS)],
                                     names=["month", "clinic_id", "account"])
    tdr = p.groupby(["month", "clinic_id", "debit"]).amount.sum()
    tcr = p.groupby(["month", "clinic_id", "credit"]).amount.sum()
    tdr.index.names = tcr.index.names = ["month", "clinic_id", "account"]
    osv = pd.DataFrame(index=idx)
    osv["turnover_debit"] = tdr.reindex(idx).fillna(0.0)
    osv["turnover_credit"] = tcr.reindex(idx).fillna(0.0)
    osv = osv.reset_index().sort_values(["clinic_id", "account", "month"])
    parts = []
    for (cid, acc), g in osv.groupby(["clinic_id", "account"]):
        g = g.copy()
        bal, opens, closes = openings.get((cid, acc), 0.0), [], []
        for t_dr, t_cr in zip(g.turnover_debit, g.turnover_credit):
            opens.append(bal)
            bal = round(bal + t_dr - t_cr, 2)
            closes.append(bal)
        g["open_net"], g["close_net"] = opens, closes
        parts.append(g)
    osv = pd.concat(parts).sort_values(["month", "clinic_id", "account"])
    osv["period"] = osv.month.map(lambda m: f"{YEAR}-{m:02d}")
    osv["account_name"] = osv.account.map(ACCOUNTS)
    osv["opening_debit"] = osv.open_net.clip(lower=0)
    osv["opening_credit"] = (-osv.open_net).clip(lower=0)
    osv["closing_debit"] = osv.close_net.clip(lower=0)
    osv["closing_credit"] = (-osv.close_net).clip(lower=0)
    cols = ["period", "clinic_id", "account", "account_name", "opening_debit", "opening_credit",
            "turnover_debit", "turnover_credit", "closing_debit", "closing_credit"]
    return osv[cols].round(2)


def rebuild_cash(post: pd.DataFrame) -> pd.DataFrame:
    c = post[(post.debit == "51") | (post.credit == "51")].copy()
    c = c[c.doc_type.isin(CASH_MAP)]
    c["activity"] = c.doc_type.map(lambda x: CASH_MAP[x][0])
    c["article"] = c.doc_type.map(lambda x: CASH_MAP[x][1])
    c["flow"] = c.doc_type.map(lambda x: CASH_MAP[x][2])
    c["signed_amount"] = np.where(c.flow == "in", c.amount, -c.amount)
    return c[["date", "clinic_id", "activity", "article", "flow", "amount",
              "signed_amount", "counterparty"]].reset_index(drop=True)

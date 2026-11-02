"""
Шаг 3 генератора: платежи ERP и бухгалтерский слой 1С за 2025 год.

Вход:  data/raw/ (clinics, services, doctors, visits, visit_items)
Выход: data/raw/
  visit_payments.csv - платежи по визитам (кассовый модуль ERP)
  postings_1c.csv    - журнал проводок (двойная запись, с аналитикой)
  osv_1c.csv         - оборотно-сальдовая ведомость по месяцам и клиникам
  cashflow_1c.csv    - движение денег по статьям ДДС
  payroll_1c.csv     - начисления врачам и персоналу
  materials_1c.csv   - плановые и фактические материалы/лаборатории

Каналы оплаты (два):
  self - пациент платит в день визита (эквайринг, комиссия 1.8%)
  dms  - ДМС: страховая платит пакетом за месяц через 30-45 дней

Принципы:
  * Данные ЧИСТЫЕ и сходятся с ERP до копейки; ошибки добавит inject_errors.py.
  * ОСВ не рисуется отдельно, а считается из журнала проводок + входящих остатков.
  * Закрытие месяца (счета 90 -> 99) отмечено doc_type='Закрытие месяца',
    отчёты P&L строятся по остальным проводкам.
  * Налог на прибыль и НДС вне рамок демо (P&L заканчивается прибылью до налога).

Запуск:  python src/generate/generate_1c.py
"""
import calendar
import random
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 7
YEAR = 2025
RAW = Path(__file__).resolve().parents[2] / "data" / "raw"
rng = random.Random(SEED)
nprng = np.random.default_rng(SEED)

# --- Параметры экономики ----------------------------------------------------
P_DMS = {"DENT": 0.12, "COSM": 0.02, "PLAS": 0.03, "FAM": 0.38, "LASER": 0.06}
INSURERS = ["СОГАЗ-Мед", "Ингосстрах", "РЕСО-Гарантия", "АльфаСтрахование", "Согласие"]
INSURER_W = [30, 25, 20, 15, 10]
ACQUIRING = 0.018            # комиссия эквайринга
NDFL = 0.13
CONTRIB = 0.22               # эффективная ставка страховых взносов (с учётом лимита)
NURSES = {"C01": 2_400_000, "C02": 1_200_000, "C03": 1_000_000}   # ФОТ медперсонала в мес.
ADMIN = {"C01": 1_000_000, "C02": 500_000, "C03": 500_000}        # ФОТ админ. персонала
DEPR = {"C01": 1_100_000, "C02": 500_000, "C03": 550_000}         # амортизация в мес.
OPEN_FA = {"C01": 70_000_000, "C02": 32_000_000, "C03": 35_000_000}  # ОС на 01.01
OPEN_CASH = {"C01": 15_000_000, "C02": 8_000_000, "C03": 8_000_000}
CAPEX = {("C01", 3): 12_000_000, ("C03", 6): 7_500_000, ("C02", 9): 4_200_000}
UTIL_PCT_OF_RENT = 0.055
MARKETING_PCT = 0.035
OPENING_LIAB_FACTOR = 1.10   # обязательства на 01.01 относительно январских

ACCOUNTS = {
    "01": "Основные средства", "02": "Амортизация основных средств", "10": "Материалы",
    "20": "Основное производство", "26": "Общехозяйственные расходы",
    "44": "Расходы на продажу", "51": "Расчётные счета",
    "60": "Расчёты с поставщиками и подрядчиками", "62.01": "Расчёты с пациентами",
    "62.02": "Расчёты со страховыми компаниями (ДМС)", "68": "Расчёты по налогам (НДФЛ)",
    "69": "Расчёты по страховым взносам", "70": "Расчёты с персоналом по оплате труда",
    "80": "Капитал (входящие остатки)", "90.01": "Выручка", "90.02": "Себестоимость продаж",
    "90.07": "Расходы на продажу (закрытие)", "90.08": "Управленческие расходы (закрытие)",
    "90.09": "Прибыль/убыток от продаж", "99": "Прибыли и убытки",
}

CASH_MAP = {  # doc_type -> (вид деятельности, статья ДДС, направление)
    "Поступление от пациентов (эквайринг)": ("operating", "Поступления от пациентов", "in"),
    "Поступление от страховой": ("operating", "Поступления от страховых компаний", "in"),
    "Комиссия эквайринга": ("operating", "Комиссия эквайринга", "out"),
    "Выплата зарплаты": ("operating", "Выплата заработной платы", "out"),
    "Уплата НДФЛ": ("operating", "НДФЛ", "out"),
    "Уплата страховых взносов": ("operating", "Страховые взносы", "out"),
    "Оплата аренды": ("operating", "Аренда помещений", "out"),
    "Оплата коммунальных услуг": ("operating", "Коммунальные услуги", "out"),
    "Оплата рекламы": ("operating", "Маркетинг и реклама", "out"),
    "Оплата поставщикам": ("operating", "Материалы и лаборатории", "out"),
    "Оплата оборудования": ("investing", "Приобретение оборудования", "out"),
}

R = []  # журнал проводок


def month_end(y, m):
    return date(y, m, calendar.monthrange(y, m)[1])


def add(d, clinic, dr, cr, amt, doc, direction="", cp="", comment=""):
    amt = round(float(amt), 2)
    if amt == 0:
        return
    R.append(dict(date=pd.Timestamp(d).date().isoformat(), clinic_id=clinic, debit=dr,
                  credit=cr, amount=amt, direction_code=direction, counterparty=cp,
                  doc_type=doc, comment=comment))


def alloc(total, shares):
    """Раскладывает сумму по ключам пропорционально весам, сумма сходится точно."""
    keys = list(shares)
    s = sum(shares.values())
    out, acc = {}, 0.0
    for k in keys[:-1]:
        out[k] = round(total * shares[k] / s, 2)
        acc += out[k]
    out[keys[-1]] = round(total - acc, 2)
    return out


def main():
    clinics = pd.read_csv(RAW / "clinics.csv")
    services = pd.read_csv(RAW / "services.csv")
    doctors = pd.read_csv(RAW / "doctors.csv")
    visits = pd.read_csv(RAW / "visits.csv", parse_dates=["visit_date"])
    items = pd.read_csv(RAW / "visit_items.csv")

    # ======================= 1. Платежи по визитам (ERP) =====================
    done = visits[visits.status == "completed"].copy()
    done = done.merge(items.groupby("visit_id").revenue.sum().rename("visit_revenue"),
                      on="visit_id", how="left")
    done["month"] = done.visit_date.dt.month
    is_dms = np.array([rng.random() < P_DMS[d] for d in done.direction_code])
    done["payer_type"] = np.where(is_dms, "dms", "self")
    done["payer_name"] = "Пациент"
    done.loc[is_dms, "payer_name"] = [rng.choices(INSURERS, INSURER_W)[0] for _ in range(is_dms.sum())]

    done["pay_date"] = done.visit_date
    keys = done.loc[is_dms, ["clinic_id", "month", "payer_name"]].drop_duplicates()
    batch = {(r.clinic_id, r.month, r.payer_name):
             pd.Timestamp(month_end(YEAR, r.month) + timedelta(days=rng.randint(30, 45)))
             for r in keys.itertuples()}
    done.loc[is_dms, "pay_date"] = [batch[(c, m, p)] for c, m, p in zip(
        done.loc[is_dms, "clinic_id"], done.loc[is_dms, "month"], done.loc[is_dms, "payer_name"])]
    year_end = pd.Timestamp(YEAR, 12, 31)

    paid = done[done.pay_date <= year_end]
    pay_df = pd.DataFrame(dict(
        payment_id=[f"PAY{i:07d}" for i in range(1, len(paid) + 1)],
        visit_id=paid.visit_id.values, pay_date=paid.pay_date.dt.date.astype(str).values,
        amount=paid.visit_revenue.values, payer_type=paid.payer_type.values,
        payer_name=paid.payer_name.values))

    im = items.merge(done[["visit_id", "visit_date", "clinic_id", "direction_code", "doctor_id",
                           "payer_name", "payer_type", "month"]], on="visit_id")
    im = im.merge(services[["service_id", "material_cost", "external_cost"]], on="service_id")
    im["mat_plan"] = im.qty * im.material_cost
    im["ext_plan"] = im.qty * im.external_cost

    # ======================= 2. Реализация и поступления =====================
    g = im.groupby(["visit_date", "clinic_id", "direction_code", "payer_name", "payer_type"]
                   ).revenue.sum().reset_index()
    for r in g.itertuples():
        dr = "62.01" if r.payer_type == "self" else "62.02"
        add(r.visit_date, r.clinic_id, dr, "90.01", r.revenue, "Реализация услуг",
            r.direction_code, "" if r.payer_type == "self" else r.payer_name)

    s = done[done.payer_type == "self"].groupby(["pay_date", "clinic_id"]).visit_revenue.sum().reset_index()
    for r in s.itertuples():
        add(r.pay_date, r.clinic_id, "51", "62.01", r.visit_revenue, "Поступление от пациентов (эквайринг)")
        add(r.pay_date, r.clinic_id, "26", "51", r.visit_revenue * ACQUIRING, "Комиссия эквайринга")

    d_ = done[(done.payer_type == "dms") & (done.pay_date <= year_end)]
    d_ = d_.groupby(["pay_date", "clinic_id", "payer_name"]).visit_revenue.sum().reset_index()
    for r in d_.itertuples():
        add(r.pay_date, r.clinic_id, "51", "62.02", r.visit_revenue, "Поступление от страховой",
            cp=r.payer_name)

    # Входящая дебиторка ДМС (декабрьские акты 2024), оплата в конце января - середине февраля
    dms_m = im[im.payer_type == "dms"].groupby(["clinic_id", "month"]).revenue.sum().reset_index()
    dms_open = {c: round(dms_m[(dms_m.clinic_id == c) & (dms_m.month >= 10)].revenue.sum() / 3 * 0.8, 2)
                for c in clinics.clinic_id}
    for c, amt in dms_open.items():
        for ins, w in zip(INSURERS, INSURER_W):
            add(date(YEAR, 1, 25) + timedelta(days=rng.randint(0, 20)), c, "51", "62.02",
                amt * w / sum(INSURER_W), "Поступление от страховой", cp=ins,
                comment="Оплата акта за декабрь 2024")

    # ======================= 3. Ежемесячные начисления =======================
    cd = im.groupby(["month", "clinic_id", "direction_code"]).agg(
        rev=("revenue", "sum"), mat=("mat_plan", "sum"), ext=("ext_plan", "sum")).reset_index()
    doc_rev = im.groupby(["doctor_id", "month"]).revenue.sum().to_dict()
    docs_by_clinic = {c: g_ for c, g_ in doctors.groupby("clinic_id")}

    liab, mat_rows, payroll_rows = {}, [], []
    for m in range(1, 13):
        last = month_end(YEAR, m)
        for c in clinics.itertuples():
            cid = c.clinic_id
            sub = cd[(cd.month == m) & (cd.clinic_id == cid)]
            if sub.empty:
                continue
            shares = dict(zip(sub.direction_code, sub.rev))
            clinic_rev = sub.rev.sum()

            # материалы и лаборатории (фактический индекс цен плавает)
            mi = 1 + 0.005 * (m - 1) + nprng.normal(0, 0.02)
            ei = 1 + 0.003 * (m - 1) + nprng.normal(0, 0.015)
            mat_total = ext_total = 0.0
            for r in sub.itertuples():
                mf, ef = round(r.mat * mi, 2), round(r.ext * ei, 2)
                add(last, cid, "20", "10", mf, "Списание материалов", r.direction_code)
                add(last, cid, "20", "60", ef, "Услуги лабораторий и сторонних организаций", r.direction_code)
                mat_rows.append(dict(period=f"{YEAR}-{m:02d}", clinic_id=cid,
                                     direction_code=r.direction_code,
                                     material_plan=round(r.mat, 2), material_fact=mf,
                                     external_plan=round(r.ext, 2), external_fact=ef,
                                     material_index=round(mi, 4), external_index=round(ei, 4)))
                mat_total += mf
                ext_total += ef
            purchase = round(mat_total * rng.uniform(0.92, 1.08), 2)
            add(last, cid, "10", "60", purchase, "Поступление материалов")

            # зарплата врачей
            gross_dir, contrib_dir = {}, {}
            ndfl_total = gross_all = contrib_all = 0
            for dd in docs_by_clinic[cid].itertuples():
                rev = doc_rev.get((dd.doctor_id, m), 0)
                if rev == 0:
                    continue
                gross = round(rev * dd.pay_pct)
                ndfl, contr = round(gross * NDFL), round(gross * CONTRIB)
                gross_dir[dd.direction_code] = gross_dir.get(dd.direction_code, 0) + gross
                contrib_dir[dd.direction_code] = contrib_dir.get(dd.direction_code, 0) + contr
                ndfl_total += ndfl
                gross_all += gross
                contrib_all += contr
                payroll_rows.append(dict(
                    period=f"{YEAR}-{m:02d}", clinic_id=cid, employee_id=dd.doctor_id,
                    employee_name=dd.doctor_name, role="doctor", direction_code=dd.direction_code,
                    base_revenue=rev, pay_pct=dd.pay_pct, gross=gross, ndfl=ndfl,
                    net_pay=gross - ndfl, contributions=contr))
            for dcode in gross_dir:
                add(last, cid, "20", "70", gross_dir[dcode], "Начисление ЗП врачей", dcode)
                add(last, cid, "20", "69", contrib_dir[dcode], "Начисление страховых взносов", dcode)

            # медперсонал: распределяется по направлениям пропорционально выручке
            nurse = int(round(NURSES[cid] * (1 + nprng.normal(0, 0.015)), -2))
            n_ndfl, n_contr = round(nurse * NDFL), round(nurse * CONTRIB)
            for dcode, v in alloc(nurse, shares).items():
                add(last, cid, "20", "70", v, "Начисление ЗП медперсонала", dcode)
            for dcode, v in alloc(n_contr, shares).items():
                add(last, cid, "20", "69", v, "Начисление страховых взносов", dcode)
            payroll_rows.append(dict(
                period=f"{YEAR}-{m:02d}", clinic_id=cid, employee_id=f"STAFF-{cid}-MED",
                employee_name=f"Средний и младший медперсонал ({cid})", role="nurse",
                direction_code="", base_revenue="", pay_pct="", gross=nurse, ndfl=n_ndfl,
                net_pay=nurse - n_ndfl, contributions=n_contr))

            # административный персонал
            adm = int(round(ADMIN[cid] * (1 + nprng.normal(0, 0.01)), -2))
            a_ndfl, a_contr = round(adm * NDFL), round(adm * CONTRIB)
            add(last, cid, "26", "70", adm, "Начисление ЗП административного персонала")
            add(last, cid, "26", "69", a_contr, "Начисление страховых взносов")
            payroll_rows.append(dict(
                period=f"{YEAR}-{m:02d}", clinic_id=cid, employee_id=f"STAFF-{cid}-ADM",
                employee_name=f"Административный персонал ({cid})", role="admin",
                direction_code="", base_revenue="", pay_pct="", gross=adm, ndfl=a_ndfl,
                net_pay=adm - a_ndfl, contributions=a_contr))

            tot_gross = gross_all + nurse + adm
            tot_ndfl = ndfl_total + n_ndfl + a_ndfl
            tot_contr = contrib_all + n_contr + a_contr
            add(last, cid, "70", "68", tot_ndfl, "Удержание НДФЛ")

            # амортизация
            for dcode, v in alloc(DEPR[cid], shares).items():
                add(last, cid, "20", "02", v, "Амортизация ОС", dcode)

            # аренда, коммунальные, маркетинг, оборудование
            rent = float(c.rent_month)
            util = round(rent * UTIL_PCT_OF_RENT * rng.uniform(0.9, 1.1), 2)
            mkt = round(clinic_rev * MARKETING_PCT * rng.uniform(0.85, 1.15), 2)
            add(date(YEAR, m, 1), cid, "26", "60", rent, "Аренда помещений")
            add(last, cid, "26", "60", util, "Коммунальные услуги")
            add(last, cid, "44", "60", mkt, "Рекламные услуги")
            capex = CAPEX.get((cid, m), 0)
            if capex:
                add(date(YEAR, m, 15), cid, "01", "60", capex, "Поступление оборудования")

            liab[(m, cid)] = dict(net=tot_gross - tot_ndfl, ndfl=tot_ndfl, contrib=tot_contr,
                                  sup=purchase + ext_total, util=util, mkt=mkt, rent=rent, equip=capex)

    # ======================= 4. Платежи (следующий месяц) ====================
    opening_liab = {c: {k: round(liab[(1, c)][k] * OPENING_LIAB_FACTOR, 2)
                        for k in ("net", "ndfl", "contrib", "sup", "util", "mkt")}
                    for c in clinics.clinic_id}
    for m in range(1, 13):
        for cid in clinics.clinic_id:
            prev = opening_liab[cid] if m == 1 else liab[(m - 1, cid)]
            add(date(YEAR, m, 5), cid, "60", "51", liab[(m, cid)]["rent"], "Оплата аренды")
            add(date(YEAR, m, 10), cid, "70", "51", prev["net"], "Выплата зарплаты")
            add(date(YEAR, m, 15), cid, "60", "51", prev["mkt"], "Оплата рекламы")
            add(date(YEAR, m, 20), cid, "60", "51", prev["sup"], "Оплата поставщикам")
            add(date(YEAR, m, 20), cid, "60", "51", prev["util"], "Оплата коммунальных услуг")
            add(date(YEAR, m, 20), cid, "60", "51", prev.get("equip", 0), "Оплата оборудования")
            add(date(YEAR, m, 28), cid, "68", "51", prev["ndfl"], "Уплата НДФЛ")
            add(date(YEAR, m, 28), cid, "69", "51", prev["contrib"], "Уплата страховых взносов")

    # ======================= 5. Закрытие месяцев (90 -> 99) ==================
    df = pd.DataFrame(R)
    df["month"] = pd.to_datetime(df.date).dt.month
    c20 = df[df.debit == "20"].groupby(["month", "clinic_id", "direction_code"]).amount.sum()
    c26 = df[df.debit == "26"].groupby(["month", "clinic_id"]).amount.sum()
    c44 = df[df.debit == "44"].groupby(["month", "clinic_id"]).amount.sum()
    rev = df[df.credit == "90.01"].groupby(["month", "clinic_id"]).amount.sum()
    for (m, cid), r_amt in rev.items():
        last = month_end(YEAR, m)
        total_cost = 0.0
        for (mm, cc, dcode), v in c20.items():
            if mm == m and cc == cid:
                add(last, cid, "90.02", "20", v, "Закрытие месяца", dcode)
                total_cost += v
        for src, acc in ((c44, "90.07"), (c26, "90.08")):
            v = src.get((m, cid), 0)
            add(last, cid, acc, "44" if acc == "90.07" else "26", v, "Закрытие месяца")
            total_cost += v
        add(last, cid, "90.01", "90.09", r_amt, "Закрытие месяца")
        add(last, cid, "90.09", "90.02", total_cost - c44.get((m, cid), 0) - c26.get((m, cid), 0), "Закрытие месяца")
        add(last, cid, "90.09", "90.07", c44.get((m, cid), 0), "Закрытие месяца")
        add(last, cid, "90.09", "90.08", c26.get((m, cid), 0), "Закрытие месяца")
        profit = round(r_amt - total_cost, 2)
        if profit >= 0:
            add(last, cid, "90.09", "99", profit, "Закрытие месяца")
        else:
            add(last, cid, "99", "90.09", -profit, "Закрытие месяца")

    # Покупка оборудования в проводках: счёт 01 (Dr) / 60 (Cr) - уже внесена выше.
    post = pd.DataFrame(R).sort_values(["date", "clinic_id", "doc_type"], kind="stable").reset_index(drop=True)
    post.insert(0, "posting_id", [f"PST{i:07d}" for i in range(1, len(post) + 1)])
    post["month"] = pd.to_datetime(post.date).dt.month

    # ======================= 6. Входящие остатки и ОСВ =======================
    openings = {}
    for cid in clinics.clinic_id:
        jan = [r for r in mat_rows if r["clinic_id"] == cid and r["period"].endswith("-01")]
        o = {"01": OPEN_FA[cid], "02": -0.35 * OPEN_FA[cid],
             "10": round(1.5 * sum(r["material_fact"] for r in jan), 2),
             "51": OPEN_CASH[cid], "62.02": dms_open[cid],
             "60": -(opening_liab[cid]["sup"] + opening_liab[cid]["util"] + opening_liab[cid]["mkt"]),
             "70": -opening_liab[cid]["net"], "68": -opening_liab[cid]["ndfl"],
             "69": -opening_liab[cid]["contrib"]}
        o["80"] = -sum(o.values())
        openings[cid] = o

    accounts = list(ACCOUNTS)
    idx = pd.MultiIndex.from_product([range(1, 13), clinics.clinic_id, accounts],
                                     names=["month", "clinic_id", "account"])
    tdr = post.groupby(["month", "clinic_id", "debit"]).amount.sum()
    tcr = post.groupby(["month", "clinic_id", "credit"]).amount.sum()
    tdr.index.names = tcr.index.names = ["month", "clinic_id", "account"]
    osv = pd.DataFrame(index=idx)
    osv["turnover_debit"] = tdr.reindex(idx).fillna(0.0)
    osv["turnover_credit"] = tcr.reindex(idx).fillna(0.0)
    osv = osv.reset_index().sort_values(["clinic_id", "account", "month"])
    osv["open_net"] = 0.0
    out = []
    for (cid, acc), grp in osv.groupby(["clinic_id", "account"]):
        grp = grp.copy()
        bal = openings[cid].get(acc, 0.0)
        opens, closes = [], []
        for t_dr, t_cr in zip(grp.turnover_debit, grp.turnover_credit):
            opens.append(bal)
            bal = round(bal + t_dr - t_cr, 2)
            closes.append(bal)
        grp["open_net"], grp["close_net"] = opens, closes
        out.append(grp)
    osv = pd.concat(out).sort_values(["month", "clinic_id", "account"])
    osv["period"] = osv.month.map(lambda m: f"{YEAR}-{m:02d}")
    osv["account_name"] = osv.account.map(ACCOUNTS)
    osv["opening_debit"] = osv.open_net.clip(lower=0).round(2)
    osv["opening_credit"] = (-osv.open_net).clip(lower=0).round(2)
    osv["closing_debit"] = osv.close_net.clip(lower=0).round(2)
    osv["closing_credit"] = (-osv.close_net).clip(lower=0).round(2)
    osv_out = osv[["period", "clinic_id", "account", "account_name", "opening_debit",
                   "opening_credit", "turnover_debit", "turnover_credit", "closing_debit",
                   "closing_credit"]].round(2)

    # --- проверки целостности ---
    chk = osv.groupby(["period", "clinic_id"])[["closing_debit", "closing_credit"]].sum()
    assert (chk.closing_debit - chk.closing_credit).abs().max() < 1, "ОСВ не сходится по сальдо"
    assert abs(post.amount.sum() * 0 + (tdr.sum() - tcr.sum())) < 1, "оборот Дт != Кт"
    rev_1c = post[(post.credit == "90.01") & (post.doc_type == "Реализация услуг")].amount.sum()
    assert abs(rev_1c - items.revenue.sum()) < 1, "выручка 1С != выручка ERP"
    assert osv[osv.account.str.startswith("90")].close_net.abs().max() < 1, "счета 90 не закрыты"

    # ======================= 7. Движение денег (ДДС) =========================
    cash = post[(post.debit == "51") | (post.credit == "51")].copy()
    cash = cash[cash.doc_type.isin(CASH_MAP)]
    cash["activity"] = cash.doc_type.map(lambda x: CASH_MAP[x][0])
    cash["article"] = cash.doc_type.map(lambda x: CASH_MAP[x][1])
    cash["flow"] = cash.doc_type.map(lambda x: CASH_MAP[x][2])
    cash["signed_amount"] = np.where(cash.flow == "in", cash.amount, -cash.amount)
    cash_out = cash[["date", "clinic_id", "activity", "article", "flow", "amount",
                     "signed_amount", "counterparty"]].reset_index(drop=True)

    # ======================= 8. Сохранение ===================================
    pay_df.to_csv(RAW / "visit_payments.csv", index=False, encoding="utf-8-sig")
    post.drop(columns="month").to_csv(RAW / "postings_1c.csv", index=False, encoding="utf-8-sig")
    osv_out.to_csv(RAW / "osv_1c.csv", index=False, encoding="utf-8-sig")
    cash_out.to_csv(RAW / "cashflow_1c.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(payroll_rows).to_csv(RAW / "payroll_1c.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(mat_rows).to_csv(RAW / "materials_1c.csv", index=False, encoding="utf-8-sig")

    # ======================= 9. Контрольная сводка ===========================
    oper = post[post.doc_type != "Закрытие месяца"]
    revenue = oper[oper.credit == "90.01"].amount.sum()
    cost20 = oper[oper.debit == "20"].amount.sum()
    cost26 = oper[oper.debit == "26"].amount.sum()
    cost44 = oper[oper.debit == "44"].amount.sum()
    print(f"Проводок: {len(post):,}   платежей ERP: {len(pay_df):,}")
    print(f"Выручка 1С = ERP: {revenue / 1e6:,.1f} млн ₽ (сходится)")
    print(f"Себестоимость (сч.20): {cost20 / 1e6:,.1f}   управленческие (26): {cost26 / 1e6:,.1f}   "
          f"коммерческие (44): {cost44 / 1e6:,.1f}")
    print(f"Валовая маржа: {(revenue - cost20) / revenue:.1%}   "
          f"Прибыль до налога: {(revenue - cost20 - cost26 - cost44) / 1e6:,.1f} млн ₽ "
          f"({(revenue - cost20 - cost26 - cost44) / revenue:.1%})")
    dec = osv[osv.period == f"{YEAR}-12"]
    print("\nОстатки на 31.12 (млн ₽), Дт(+)/Кт(-):")
    piv = dec.pivot(index="account", columns="clinic_id", values="close_net") / 1e6
    print(piv.loc[["51", "62.02", "60", "70", "68", "69", "10", "99"]].round(1).to_string())
    cash51 = osv[osv.account == "51"].groupby("clinic_id").close_net.min() / 1e6
    print("\nМинимальный остаток на р/с за год (млн ₽):")
    print(cash51.round(1).to_string())
    print("\nДДС за год по статьям (млн ₽):")
    print((cash_out.groupby("article").signed_amount.sum() / 1e6).round(1).to_string())


if __name__ == "__main__":
    main()

"""
Шаг 5 генератора: внесение ошибок в данные для модуля сверки.

Чистые файлы в data/raw/ остаются эталоном. «Испорченные» копии пишутся
в data/dirty/ (их и анализирует модуль сверки). Ключ ответов - в
data/answer_key/injected_errors.csv (агенту не показывать).

14 ошибок (по типам):
  E01-E03  выручка 1С != выручка ERP (ручная правка реализации в 1С)
  E04-E05  дубли визитов в ERP
  E06-E07  визит закрыт, оплаты нет (платёж по самоплате потерян)
  E08-E09  визит без врача
  E10-E11  цена ниже прайса без скидки
  E12-E13  выручка 1С отнесена не в тот период (закрыта следующим месяцем)
  E14      задвоенное поступление на р/с -> кредитовое сальдо на 62.01
Плюс 12 «безобидных» расхождений в 1-40 руб. (округления). Это НЕ ошибки:
модуль сверки должен их пропускать (проверка на ложные срабатывания).

Принципы:
  * все правки в 1С делаются на проводках, затем пересчитываются закрытие
    месяца, ОСВ и ДДС, поэтому файлы остаются внутренне согласованными;
  * ERP-правки выполняются в visits / visit_items / visit_payments;
  * ID записей в ключе соответствуют ID в испорченных файлах.

Запуск:  python src/generate/inject_errors.py
"""
import calendar
import random
import shutil
from datetime import date, timedelta
from pathlib import Path

import sys

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.ledger import rebuild_closing, rebuild_cash, rebuild_osv  # noqa: E402


SEED = 21
YEAR = 2025
RAW = Path(__file__).resolve().parents[2] / "data" / "raw"
DIRTY = RAW.parent / "dirty"
KEY = RAW.parent / "answer_key"
rng = random.Random(SEED)

STR = {"debit": str, "credit": str, "account": str}


def month_end(y, m):
    return date(y, m, calendar.monthrange(y, m)[1])


def main():
    DIRTY.mkdir(parents=True, exist_ok=True)
    KEY.mkdir(parents=True, exist_ok=True)

    visits = pd.read_csv(RAW / "visits.csv")
    items = pd.read_csv(RAW / "visit_items.csv")
    pays = pd.read_csv(RAW / "visit_payments.csv")
    post = pd.read_csv(RAW / "postings_1c.csv", dtype=STR)
    osv_clean = pd.read_csv(RAW / "osv_1c.csv", dtype=STR | {c: float for c in [
        "opening_debit", "opening_credit", "turnover_debit", "turnover_credit",
        "closing_debit", "closing_credit"]})

    post["_tag"] = ""
    oper = post[post.doc_type != "Закрытие месяца"].copy()
    oper["month"] = pd.to_datetime(oper.date).dt.month
    vm = visits.copy()
    vm["month"] = pd.to_datetime(vm.visit_date).dt.month
    vrev = items.groupby("visit_id").revenue.sum()
    vm["rev"] = vm.visit_id.map(vrev).fillna(0)
    key_rows = []
    used_visits, used_posts = set(), set()

    def note(**kw):
        key_rows.append(kw)

    # ---------- E01-E03: выручка 1С != ERP (ручная правка реализации) --------
    for eid, cid, m, delta in [("E01", "C01", 2, 212_400), ("E02", "C03", 6, 96_700),
                               ("E03", "C02", 10, -143_500)]:
        cand = oper[(oper.clinic_id == cid) & (oper.month == m) & (oper.credit == "90.01") &
                    (oper.debit == "62.01") & (oper.amount > 2 * abs(delta))]
        i = rng.choice(list(cand.index))
        oper.loc[i, "amount"] = round(oper.loc[i, "amount"] + delta, 2)
        oper.loc[i, "_tag"] = eid
        used_posts.add(i)
        note(error_id=eid, error_type="revenue_mismatch", file="postings_1c.csv", period=f"{YEAR}-{m:02d}",
             clinic_id=cid, direction_code=oper.loc[i, "direction_code"], amount_effect_rub=delta,
             description=f"Реализация в 1С за {oper.loc[i, 'date']} изменена на {delta:+,} ₽ относительно ERP; "
                         f"симптом: расхождение выручки и остаток на 62.01")

    # ---------- E04-E05: дубли визитов в ERP --------------------------------
    next_v = int(vm.visit_id.str[1:].astype(int).max()) + 1
    next_i = int(items.item_id.str[1:].astype(int).max()) + 1
    new_visits, new_items = [], []
    for eid, cid, m in [("E04", "C01", 4), ("E05", "C03", 11)]:
        cand = vm[(vm.clinic_id == cid) & (vm.month == m) & (vm.status == "completed") & (vm.rev >= 60_000)]
        src = cand.loc[rng.choice(list(cand.index))]
        used_visits.add(src.visit_id)
        nv = f"V{next_v:07d}"
        next_v += 1
        row = src.drop(labels=["month", "rev"]).to_dict()
        row["visit_id"] = nv
        new_visits.append(row)
        for it in items[items.visit_id == src.visit_id].to_dict("records"):
            it = dict(it, item_id=f"I{next_i:07d}", visit_id=nv)
            next_i += 1
            new_items.append(it)
        note(error_id=eid, error_type="duplicate_visit", file="visits.csv/visit_items.csv",
             record_ids=f"{src.visit_id}|{nv}", period=f"{YEAR}-{m:02d}", clinic_id=cid,
             direction_code=src.direction_code, amount_effect_rub=float(src.rev),
             description=f"Визит {src.visit_id} задвоен ({nv}): тот же пациент, врач, дата и время; "
                         f"выручка ERP завышена; у дубля нет платежа")

    # ---------- E06-E07: визит закрыт, оплаты нет ---------------------------
    pv = pays.merge(vm[["visit_id", "clinic_id", "month", "rev"]], on="visit_id")
    for eid, cid, m in [("E06", "C02", 6), ("E07", "C01", 9)]:
        cand = pv[(pv.clinic_id == cid) & (pv.month == m) & (pv.payer_type == "self") &
                  (pv.amount.between(30_000, 150_000)) & (~pv.visit_id.isin(used_visits))]
        r = cand.loc[rng.choice(list(cand.index))]
        used_visits.add(r.visit_id)
        pays = pays[pays.payment_id != r.payment_id]
        note(error_id=eid, error_type="no_payment", file="visit_payments.csv", record_ids=r.visit_id,
             period=f"{YEAR}-{m:02d}", clinic_id=cid,
             direction_code=vm.loc[vm.visit_id == r.visit_id, "direction_code"].iloc[0],
             amount_effect_rub=float(r.amount),
             description=f"Самоплата по визиту {r.visit_id} не найдена в платежах ERP, "
                         f"хотя деньги поступили в 1С в день визита")

    # ---------- E08-E09: визит без врача ------------------------------------
    for eid, cid, m in [("E08", "C03", 3), ("E09", "C02", 8)]:
        cand = vm[(vm.clinic_id == cid) & (vm.month == m) & (vm.status == "completed") &
                  (vm.rev >= 50_000) & (~vm.visit_id.isin(used_visits))]
        r = cand.loc[rng.choice(list(cand.index))]
        used_visits.add(r.visit_id)
        visits.loc[visits.visit_id == r.visit_id, "doctor_id"] = np.nan
        note(error_id=eid, error_type="no_doctor", file="visits.csv", record_ids=r.visit_id,
             period=f"{YEAR}-{m:02d}", clinic_id=cid, direction_code=r.direction_code,
             amount_effect_rub=float(r.rev),
             description=f"У завершённого визита {r.visit_id} не указан врач: выручка не атрибутирована, "
                         f"расчёт мотивации врача искажён")

    # ---------- E10-E11: цена ниже прайса без скидки --------------------------
    im = items.merge(vm[["visit_id", "clinic_id", "month"]], on="visit_id")
    for eid, cid, m in [("E10", "C01", 7), ("E11", "C03", 5)]:
        cand = im[(im.clinic_id == cid) & (im.month == m) & (im.discount_pct == 0) &
                  (im.list_price_at_date >= 60_000) & (~im.visit_id.isin(used_visits))]
        r = cand.loc[rng.choice(list(cand.index))]
        used_visits.add(r.visit_id)
        factor = 0.85 if eid == "E10" else 0.80
        new_price = int(round(r.list_price_at_date * factor / 100) * 100)
        old_rev = float(r.revenue)
        items.loc[items.item_id == r.item_id, ["price_fact", "revenue"]] = [new_price, new_price * int(r.qty)]
        note(error_id=eid, error_type="price_below_list", file="visit_items.csv", record_ids=r.item_id,
             period=f"{YEAR}-{m:02d}", clinic_id=cid,
             direction_code=visits.loc[visits.visit_id == r.visit_id, "direction_code"].iloc[0],
             amount_effect_rub=float(new_price * int(r.qty)) - old_rev,
             description=f"Позиция {r.item_id}: фактическая цена {new_price:,} при прайсе "
                         f"{int(r.list_price_at_date):,} и нулевой скидке")

    # ---------- E12-E13: выручка 1С в неверном периоде -----------------------
    for eid, cid, m in [("E12", "C01", 3), ("E13", "C02", 8)]:
        sub = oper[(oper.clinic_id == cid) & (oper.month == m) & (oper.credit == "90.01") &
                   (oper.debit == "62.01") & (~oper.index.isin(used_posts))]
        last_day = sub.date.max()
        cand = sub[sub.date == last_day]
        i = cand.amount.idxmax()
        used_posts.add(i)
        new_date = (month_end(YEAR, m) + timedelta(days=1)).isoformat()
        old_date = oper.loc[i, "date"]
        oper.loc[i, "date"] = new_date
        oper.loc[i, "_tag"] = eid
        note(error_id=eid, error_type="wrong_period", file="postings_1c.csv", period=f"{YEAR}-{m:02d}",
             clinic_id=cid, direction_code=oper.loc[i, "direction_code"],
             amount_effect_rub=float(oper.loc[i, "amount"]),
             description=f"Реализация {old_date} проведена {new_date}: выручка {m:02d} занижена, "
                         f"следующего месяца завышена")

    # ---------- E14: задвоенное поступление на р/с ---------------------------
    cand = oper[(oper.clinic_id == "C02") & (oper.doc_type == "Поступление от пациентов (эквайринг)") &
                (oper.date.between(f"{YEAR}-09-10", f"{YEAR}-09-20"))]
    i = cand.amount.idxmax()
    dup = oper.loc[[i]].copy()
    dup["_tag"] = "E14"
    oper = pd.concat([oper, dup], ignore_index=True)
    note(error_id="E14", error_type="negative_balance", file="postings_1c.csv/cashflow_1c.csv",
         period=f"{YEAR}-09", clinic_id="C02", direction_code="",
         amount_effect_rub=float(dup.amount.iloc[0]),
         description=f"Поступление от {dup.date.iloc[0]} проведено дважды: кредитовое сальдо на 62.01, "
                     f"ДДС завышен, в платежах ERP пары нет")

    # ---------- безобидные округления (НЕ ошибки) ----------------------------
    benign_pool = oper[(oper.debit == "62.02") & (oper.credit == "90.01") & (oper._tag == "")]
    for i in rng.sample(list(benign_pool.index), 12):
        d = rng.choice([-1, 1]) * rng.randint(1, 40)
        oper.loc[i, "amount"] = round(oper.loc[i, "amount"] + d, 2)
        oper.loc[i, "_tag"] = "BENIGN"
    for i in oper.index[oper._tag == "BENIGN"]:
        note(error_id="BENIGN", error_type="rounding_noise", file="postings_1c.csv",
             period=oper.loc[i, "date"][:7], clinic_id=oper.loc[i, "clinic_id"],
             direction_code=oper.loc[i, "direction_code"], amount_effect_rub=np.nan,
             description="Расхождение в рублях (округление): НЕ ошибка, сверка должна пропустить")

    # ============================ сборка и запись ============================
    oper = oper.drop(columns="month")
    closing = rebuild_closing(oper)
    dirty_post = pd.concat([oper, closing], ignore_index=True)
    dirty_post = dirty_post.sort_values(["date", "clinic_id", "doc_type"], kind="stable").reset_index(drop=True)
    dirty_post["posting_id"] = [f"PST{i:07d}" for i in range(1, len(dirty_post) + 1)]

    dirty_visits = pd.concat([visits, pd.DataFrame(new_visits)], ignore_index=True)
    dirty_items = pd.concat([items, pd.DataFrame(new_items)], ignore_index=True)

    osv_dirty = rebuild_osv(dirty_post.drop(columns="_tag"), osv_clean)
    chk = osv_dirty.groupby(["period", "clinic_id"])[["closing_debit", "closing_credit"]].sum()
    assert (chk.closing_debit - chk.closing_credit).abs().max() < 1, "ОСВ не сходится"
    cash_dirty = rebuild_cash(dirty_post)

    # id проводок в ключ
    for r in key_rows:
        tag = r["error_id"]
        if tag in ("BENIGN",):
            continue
        if r["file"].startswith("postings_1c") or tag == "E14":
            ids = dirty_post.loc[dirty_post._tag == tag, "posting_id"].tolist()
            r["record_ids"] = "|".join(ids)

    clean_copy = ["clinics", "doctors", "services", "patients", "price_history", "payroll_1c",
                  "materials_1c", "budget_revenue", "budget_opex"]
    for name in clean_copy:
        shutil.copy(RAW / f"{name}.csv", DIRTY / f"{name}.csv")
    enc = dict(index=False, encoding="utf-8-sig")
    dirty_visits.to_csv(DIRTY / "visits.csv", **enc)
    dirty_items.to_csv(DIRTY / "visit_items.csv", **enc)
    pays.to_csv(DIRTY / "visit_payments.csv", **enc)
    dirty_post.drop(columns="_tag")[post.columns.drop("_tag")].to_csv(DIRTY / "postings_1c.csv", **enc)
    osv_dirty.to_csv(DIRTY / "osv_1c.csv", **enc)
    cash_dirty.to_csv(DIRTY / "cashflow_1c.csv", **enc)

    key = pd.DataFrame(key_rows)
    cols = ["error_id", "error_type", "file", "record_ids", "period", "clinic_id", "direction_code",
            "amount_effect_rub", "description"]
    key[cols].to_csv(KEY / "injected_errors.csv", **enc)

    # ============================ сводка =====================================
    errs = key[key.error_id != "BENIGN"]
    print(f"Внесено ошибок: {len(errs)}  безобидных округлений: {(key.error_id == 'BENIGN').sum()}")
    print(errs.error_type.value_counts().to_string())

    vd = dirty_visits.assign(month=pd.to_datetime(dirty_visits.visit_date).dt.month)
    erp = dirty_items.merge(vd[["visit_id", "month"]], on="visit_id").groupby("month").revenue.sum()
    onec = dirty_post[(dirty_post.credit == "90.01") & (dirty_post.doc_type == "Реализация услуг")]
    onec = onec.assign(month=pd.to_datetime(onec.date).dt.month).groupby("month").amount.sum()
    diff = (onec - erp).round(0)
    print("\nРасхождение выручки 1С - ERP по месяцам (руб.), было 0:")
    print(diff[diff.abs() > 100].to_string())
    neg = osv_dirty[(osv_dirty.account == "62.01") & (osv_dirty.closing_credit > 100)]
    print("\nКредитовое сальдо 62.01 на конец периода:")
    print(neg[["period", "clinic_id", "closing_credit"]].to_string(index=False))


if __name__ == "__main__":
    main()

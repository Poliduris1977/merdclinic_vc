"""
Шаг 5b генератора: ВТОРАЯ версия внедрения ошибок (проверка обобщаемости сверки).

Сверка писалась и отлаживалась на ошибках из inject_errors.py (сид 21), поэтому её 100% полноты
там - регрессионный тест. Здесь другой сид (77) и три группы ошибок:

  A. Известные типы в новых местах и с новыми суммами (проверка, что правила не привязаны к конкретным случаям):
     V01 дубль визита, V02 нет оплаты, V03 цена ниже прайса, V04 ручная правка реализации в 1С,
     V05 задвоенное поступление на р/с.
  B. Типы, для которых в reconcile.py правила ЕСТЬ, но в первой версии теста ошибок такого типа не было:
     V06 платежи ERP не проведены в 1С (потерянное поступление), V07 поступление от страховой на другую сумму.
  C. Типы, для которых правил НЕТ (проверка на неизвестное):
     V08-V09 врач из другой клиники в визите (ERP),
     V10     реализация в 1С проведена не на ту клинику,
     V11     «фантомная» реализация в 1С без визита в ERP,
     V12     опечатка в количестве: ERP и 1С согласованно завышены (расхождения между системами нет),
     V13     двойное бронирование врача (один врач, одно время, два пациента).

Плюс 10 «безобидных» округлений в 1-40 руб. (НЕ ошибки).

Выход: data/dirty_v2/ (испорченные копии) и data/answer_key_v2/injected_errors.csv (ключ, агенту не показывать).
Чистые файлы data/raw/ и первая версия (data/dirty/, data/answer_key/) не затрагиваются.

Запуск:  python src/generate/inject_errors_v2.py
"""
import random
import shutil
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.ledger import rebuild_closing, rebuild_cash, rebuild_osv  # noqa: E402

SEED = 77
YEAR = 2025
RAW = Path(__file__).resolve().parents[2] / "data" / "raw"
DIRTY = RAW.parent / "dirty_v2"
KEY = RAW.parent / "answer_key_v2"
rng = random.Random(SEED)
STR = {"debit": str, "credit": str, "account": str}
REAL = "Реализация услуг"
RCPT = "Поступление от пациентов (эквайринг)"
DMS = "Поступление от страховой"


def main():
    DIRTY.mkdir(parents=True, exist_ok=True)
    KEY.mkdir(parents=True, exist_ok=True)

    visits = pd.read_csv(RAW / "visits.csv")
    items = pd.read_csv(RAW / "visit_items.csv")
    pays = pd.read_csv(RAW / "visit_payments.csv")
    doctors = pd.read_csv(RAW / "doctors.csv")
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
    key_rows, used_visits, used_posts = [], set(), set()
    deleted = []

    def note(**kw):
        key_rows.append(kw)

    def pick_post(mask):
        cand = oper[mask & ~oper.index.isin(used_posts)]
        i = rng.choice(list(cand.index))
        used_posts.add(i)
        return i

    # ---------- A. известные типы в новых местах ------------------------------------------------
    # V01: дубль визита (C02, февраль)
    next_v = int(vm.visit_id.str[1:].astype(int).max()) + 1
    next_i = int(items.item_id.str[1:].astype(int).max()) + 1
    new_visits, new_items = [], []
    cand = vm[(vm.clinic_id == "C02") & (vm.month == 2) & (vm.status == "completed") & (vm.rev >= 40_000)]
    src = cand.loc[rng.choice(list(cand.index))]
    used_visits.add(src.visit_id)
    nv = f"V{next_v:07d}"
    row = src.drop(labels=["month", "rev"]).to_dict()
    row["visit_id"] = nv
    new_visits.append(row)
    for it in items[items.visit_id == src.visit_id].to_dict("records"):
        new_items.append(dict(it, item_id=f"I{next_i:07d}", visit_id=nv))
        next_i += 1
    note(error_id="V01", error_type="duplicate_visit", file="visits.csv/visit_items.csv",
         record_ids=f"{src.visit_id}|{nv}", period=f"{YEAR}-02", clinic_id="C02",
         direction_code=src.direction_code, amount_effect_rub=float(src.rev),
         description=f"Визит {src.visit_id} задвоен ({nv}): выручка ERP завышена, у дубля нет платежа")

    # V02: нет оплаты (C03, май)
    pv = pays.merge(vm[["visit_id", "clinic_id", "month"]], on="visit_id")
    cand = pv[(pv.clinic_id == "C03") & (pv.month == 5) & (pv.payer_type == "self") &
              (pv.amount.between(20_000, 90_000)) & (~pv.visit_id.isin(used_visits))]
    r = cand.loc[rng.choice(list(cand.index))]
    used_visits.add(r.visit_id)
    pays = pays[pays.payment_id != r.payment_id]
    note(error_id="V02", error_type="no_payment", file="visit_payments.csv", record_ids=r.visit_id,
         period=f"{YEAR}-05", clinic_id="C03",
         direction_code=vm.loc[vm.visit_id == r.visit_id, "direction_code"].iloc[0],
         amount_effect_rub=float(r.amount),
         description=f"Самоплата по визиту {r.visit_id} отсутствует в платежах ERP, деньги поступили в 1С")

    # V03: цена ниже прайса без скидки (C02, октябрь)
    im = items.merge(vm[["visit_id", "clinic_id", "month"]], on="visit_id")
    cand = im[(im.clinic_id == "C02") & (im.month == 10) & (im.discount_pct == 0) &
              (im.list_price_at_date >= 40_000) & (~im.visit_id.isin(used_visits))]
    r = cand.loc[rng.choice(list(cand.index))]
    used_visits.add(r.visit_id)
    new_price = int(round(r.list_price_at_date * 0.90 / 100) * 100)
    old_rev = float(r.revenue)
    items.loc[items.item_id == r.item_id, ["price_fact", "revenue"]] = [new_price, new_price * int(r.qty)]
    note(error_id="V03", error_type="price_below_list", file="visit_items.csv", record_ids=r.item_id,
         period=f"{YEAR}-10", clinic_id="C02",
         direction_code=vm.loc[vm.visit_id == r.visit_id, "direction_code"].iloc[0],
         amount_effect_rub=float(new_price * int(r.qty)) - old_rev,
         description=f"Позиция {r.item_id}: цена {new_price:,} при прайсе {int(r.list_price_at_date):,}, скидка 0")

    # V04: ручная правка реализации в 1С (C01, август)
    i = pick_post((oper.clinic_id == "C01") & (oper.month == 8) & (oper.doc_type == REAL) &
                  (oper.credit == "90.01") & (oper.debit == "62.01") & (oper.amount > 60_000))
    delta = 178_300
    oper.loc[i, "amount"] = round(oper.loc[i, "amount"] + delta, 2)
    oper.loc[i, "_tag"] = "V04"
    note(error_id="V04", error_type="revenue_mismatch", file="postings_1c.csv", period=f"{YEAR}-08",
         clinic_id="C01", direction_code=oper.loc[i, "direction_code"], amount_effect_rub=delta,
         description=f"Реализация в 1С за {oper.loc[i, 'date']} изменена на {delta:+,} ₽ относительно ERP")

    # V05: задвоенное поступление на р/с (C03, август)
    cand = oper[(oper.clinic_id == "C03") & (oper.doc_type == RCPT) & (oper.date.between(f"{YEAR}-08-05", f"{YEAR}-08-25"))
                & ~oper.index.isin(used_posts)]
    i = cand.amount.idxmax()
    used_posts.add(i)
    dup = oper.loc[[i]].copy()
    dup["_tag"] = "V05"
    oper = pd.concat([oper, dup], ignore_index=True)
    note(error_id="V05", error_type="negative_balance", file="postings_1c.csv/cashflow_1c.csv",
         period=f"{YEAR}-08", clinic_id="C03", direction_code="", amount_effect_rub=float(dup.amount.iloc[0]),
         description=f"Поступление от {dup.date.iloc[0]} проведено дважды")

    # ---------- B. правила есть, ошибок такого типа раньше не было -----------------------------
    # V06: платежи ERP не проведены в 1С (потерянное поступление), C03, март
    i = pick_post((oper.clinic_id == "C03") & (oper.month == 3) & (oper.doc_type == RCPT) &
                  oper.amount.between(100_000, 400_000))
    lost = float(oper.loc[i, "amount"])
    lost_date, lost_id = oper.loc[i, "date"], oper.loc[i, "posting_id"]
    oper = oper.drop(index=i)
    note(error_id="V06", error_type="missing_receipt", file="postings_1c.csv", record_ids="",
         period=f"{YEAR}-03", clinic_id="C03", direction_code="", amount_effect_rub=-lost,
         description=f"Поступление от пациентов за {lost_date} ({lost:,.0f} ₽) не проведено в 1С; "
                     f"платежи в ERP есть (исходная проводка {lost_id})")

    # V07: поступление от страховой на другую сумму (C01)
    i = pick_post((oper.clinic_id == "C01") & (oper.doc_type == DMS) & (oper.amount > 300_000) &
                  (oper.comment.fillna("") != "Оплата акта за декабрь 2024") & (oper.month.between(3, 9)))
    oper.loc[i, "amount"] = round(oper.loc[i, "amount"] - 48_000, 2)
    oper.loc[i, "_tag"] = "V07"
    note(error_id="V07", error_type="dms_receipt_mismatch", file="postings_1c.csv", record_ids="",
         period=oper.loc[i, "date"][:7], clinic_id="C01", direction_code="", amount_effect_rub=-48_000,
         description=f"Поступление от страховой {oper.loc[i, 'counterparty']} за {oper.loc[i, 'date']} "
                     f"проведено на 48 000 ₽ меньше платежа в ERP")

    # ---------- C. новые типы, правил нет -------------------------------------------------------
    # V08-V09: врач из другой клиники в визите
    for eid, cid, m, dcode in (("V08", "C01", 4, "DENT"), ("V09", "C02", 9, "FAM")):
        cand = vm[(vm.clinic_id == cid) & (vm.month == m) & (vm.direction_code == dcode) &
                  (vm.status == "completed") & vm.doctor_id.notna() & (vm.rev >= 5_000) &
                  ~vm.visit_id.isin(used_visits)]
        r = cand.loc[rng.choice(list(cand.index))]
        other = doctors[(doctors.direction_code == dcode) & (doctors.clinic_id != cid)]
        assert len(other), f"нет врачей {dcode} в других клиниках"
        new_doc = other.iloc[rng.randrange(len(other))]
        used_visits.add(r.visit_id)
        visits.loc[visits.visit_id == r.visit_id, "doctor_id"] = new_doc.doctor_id
        note(error_id=eid, error_type="doctor_wrong_clinic", file="visits.csv", record_ids=r.visit_id,
             period=f"{YEAR}-{m:02d}", clinic_id=cid, direction_code=dcode, amount_effect_rub=float(r.rev),
             description=f"В визите {r.visit_id} клиники {cid} указан врач {new_doc.doctor_id} клиники "
                         f"{new_doc.clinic_id}: выручка попадёт в мотивацию чужого врача")

    # V10: реализация в 1С проведена не на ту клинику
    i = pick_post((oper.clinic_id == "C01") & (oper.month == 6) & (oper.doc_type == REAL) &
                  (oper.credit == "90.01") & (oper.debit == "62.01") &
                  (oper.direction_code == "DENT") & oper.amount.between(40_000, 200_000))
    amt = float(oper.loc[i, "amount"])
    oper.loc[i, "clinic_id"] = "C02"
    oper.loc[i, "_tag"] = "V10"
    note(error_id="V10", error_type="posting_wrong_clinic", file="postings_1c.csv", period=f"{YEAR}-06",
         clinic_id="C01", direction_code="DENT", amount_effect_rub=amt,
         description=f"Реализация {oper.loc[i, 'date']} DENT на {amt:,.0f} ₽ проведена на клинику C02 вместо C01: "
                     f"в C01 выручка занижена, в C02 завышена")

    # V11: «фантомная» реализация в 1С без визита в ERP
    phantom = dict(posting_id="PHANTOM", date=f"{YEAR}-11-14", clinic_id="C02", debit="62.01", credit="90.01",
                   amount=134_800.0, direction_code="COSM", counterparty=np.nan, doc_type=REAL,
                   comment=np.nan, _tag="V11", month=11)
    oper = pd.concat([oper, pd.DataFrame([phantom])], ignore_index=True)
    note(error_id="V11", error_type="phantom_revenue", file="postings_1c.csv", period=f"{YEAR}-11",
         clinic_id="C02", direction_code="COSM", amount_effect_rub=134_800.0,
         description="Реализация на 134 800 ₽ проведена в 1С вручную, визитов-оснований в ERP нет")

    # V12: опечатка в количестве, ERP и 1С согласованно завышены
    pv = pays.merge(vm[["visit_id", "clinic_id", "month", "direction_code"]], on="visit_id")
    cand_items = im[(im.clinic_id == "C03") & (im.month == 3) & (im.qty == 1) &
                    (im.price_fact.between(8_000, 20_000)) & (~im.visit_id.isin(used_visits))]
    cand_items = cand_items[cand_items.visit_id.isin(pv[pv.payer_type == "self"].visit_id)]
    r = cand_items.loc[rng.choice(list(cand_items.index))]
    used_visits.add(r.visit_id)
    pay = pays[pays.visit_id == r.visit_id].iloc[0]
    assert abs(pay.amount - vm.loc[vm.visit_id == r.visit_id, "rev"].iloc[0]) < 1, "оплата != выручка визита"
    delta = float(r.price_fact * 9)
    items.loc[items.item_id == r.item_id, ["qty", "revenue"]] = [10, r.price_fact * 10]
    pays.loc[pays.payment_id == pay.payment_id, "amount"] = pay.amount + delta
    vdate = vm.loc[vm.visit_id == r.visit_id, "visit_date"].iloc[0]
    ddir = vm.loc[vm.visit_id == r.visit_id, "direction_code"].iloc[0]
    j = oper[(oper.clinic_id == "C03") & (oper.date == vdate) & (oper.doc_type == REAL) & (oper.debit == "62.01") &
             (oper.direction_code == ddir)].index
    k = oper[(oper.clinic_id == "C03") & (oper.date == pay.pay_date) & (oper.doc_type == RCPT)].index
    assert len(j) == 1 and len(k) == 1, "не нашли проводки для согласованной правки"
    oper.loc[j[0], "amount"] = round(oper.loc[j[0], "amount"] + delta, 2)
    oper.loc[k[0], "amount"] = round(oper.loc[k[0], "amount"] + delta, 2)
    note(error_id="V12", error_type="qty_typo_consistent", file="visit_items.csv", record_ids=r.item_id,
         period=f"{YEAR}-03", clinic_id="C03", direction_code=ddir, amount_effect_rub=delta,
         description=f"Позиция {r.item_id}: количество 10 вместо 1 при цене {int(r.price_fact):,}. Ошибка проведена "
                     f"согласованно в ERP (позиция, платёж) и в 1С (реализация, поступление): расхождения "
                     f"между системами нет, виден только выброс по количеству")

    # V13: двойное бронирование врача (один врач, дата и время, два разных пациента)
    cand = vm[(vm.clinic_id == "C02") & (vm.month == 11) & (vm.status == "completed") & vm.doctor_id.notna() &
              ~vm.visit_id.isin(used_visits)]
    done = False
    for (doc, dt), g in cand.groupby(["doctor_id", "visit_date"]):
        g = g.sort_values("visit_time")
        if len(g) >= 2 and g.visit_time.nunique() >= 2 and g.patient_id.nunique() >= 2:
            a, b = g.iloc[0], g.iloc[1]
            visits.loc[visits.visit_id == b.visit_id, "visit_time"] = a.visit_time
            used_visits |= {a.visit_id, b.visit_id}
            note(error_id="V13", error_type="double_booked_doctor", file="visits.csv",
                 record_ids=f"{a.visit_id}|{b.visit_id}", period=f"{YEAR}-11", clinic_id="C02",
                 direction_code=a.direction_code, amount_effect_rub=np.nan,
                 description=f"Врач {doc} {dt} в {a.visit_time} принимает двух пациентов одновременно "
                             f"({a.visit_id}, {b.visit_id})")
            done = True
            break
    assert done, "не нашли подходящую пару визитов для V13"

    # ---------- безобидные округления (НЕ ошибки) ----------------------------------------------
    benign_pool = oper[(oper.debit == "62.02") & (oper.credit == "90.01") & (oper._tag == "")]
    for i in rng.sample(list(benign_pool.index), 10):
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

    for r in key_rows:
        tag = r["error_id"]
        if tag in ("BENIGN", "V06"):
            continue
        if r["file"].startswith("postings_1c"):
            ids = dirty_post.loc[dirty_post._tag == tag, "posting_id"].tolist()
            r["record_ids"] = "|".join(ids)

    for name in ["clinics", "doctors", "services", "patients", "price_history", "payroll_1c",
                 "materials_1c", "budget_revenue", "budget_opex"]:
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

    errs = key[key.error_id != "BENIGN"]
    print(f"Внесено ошибок: {len(errs)}  безобидных округлений: {(key.error_id == 'BENIGN').sum()}")
    print(errs.error_type.value_counts().to_string())


if __name__ == "__main__":
    main()

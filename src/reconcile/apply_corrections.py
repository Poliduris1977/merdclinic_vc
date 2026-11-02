"""
Шаг 7: исправление данных по находкам сверки.

Вход:  data/dirty/ (данные с ошибками) + data/output/findings.csv (находки сверки)
Выход: data/corrected/        - исправленный набор данных (все файлы)
       data/output/corrections_log.csv - журнал: что, где, до/после, на каком основании
       data/output/unresolved.csv      - то, что автоматически исправить нельзя

Ключ ответов НЕ используется. Каждое исправление имеет основание и уровень
уверенности; спорные случаи помечены для подтверждения бухгалтером.

Что делается по типам находок:
  duplicate_visit    - удаляется дубль (визит и его услуги)             [уверенность: высокая]
  price_below_list   - цена приводится к прайсу с учётом скидки         [высокая, если платёж = цене по прайсу]
  no_payment         - восстанавливается платёж в ERP на сумму визита    [высокая: деньги есть в 1С в день визита]
  revenue_mismatch   - корректируется проводка реализации в 1С          [средняя: проводка выбрана из кандидатов]
  wrong_period       - проводка возвращается на дату визитов ERP         [высокая]
  duplicate_receipt  - удаляется лишнее поступление на р/с              [высокая]
  no_doctor          - НЕ исправляется (врача восстановить нельзя)       [ручной разбор]
После правок закрытие месяца, ОСВ и ДДС пересчитываются.

Запуск:  python src/reconcile/apply_corrections.py
"""
import shutil
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common.ledger import rebuild_cash, rebuild_closing, rebuild_osv  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
DIRTY = ROOT / "data" / "dirty"
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "output"
TOL = 100.0
STR = {"debit": str, "credit": str, "account": str}
NUM = ["opening_debit", "opening_credit", "turnover_debit", "turnover_credit",
       "closing_debit", "closing_credit"]


def main():
    CORR.mkdir(parents=True, exist_ok=True)
    visits = pd.read_csv(DIRTY / "visits.csv")
    items = pd.read_csv(DIRTY / "visit_items.csv")
    pays = pd.read_csv(DIRTY / "visit_payments.csv")
    post = pd.read_csv(DIRTY / "postings_1c.csv", dtype=STR)
    osv_dirty = pd.read_csv(DIRTY / "osv_1c.csv", dtype=STR | {c: float for c in NUM})
    findings = pd.read_csv(OUT / "findings.csv")
    roots = findings[findings.is_root_cause]

    log, manual = [], []

    def record(fid, etype, action, target, before, after, basis, confidence, status="applied"):
        log.append(dict(finding_id=fid, error_type=etype, action=action, target=target,
                        before=before, after=after, basis=basis, confidence=confidence, status=status))

    rev_before = items.revenue.sum()
    next_pay = int(pays.payment_id.str[3:].astype(int).max()) + 1

    for f in roots.itertuples():
        ids = str(f.record_ids).split("|")

        # ---- дубль визита ----
        if f.error_type == "duplicate_visit":
            orig, dup = ids
            amt = items.loc[items.visit_id == dup, "revenue"].sum()
            has_pay = dup in set(pays.visit_id)
            visits = visits[visits.visit_id != dup]
            items = items[items.visit_id != dup]
            record(f.finding_id, f.error_type, "удалить визит и услуги", dup, f"выручка {amt:,.0f}", "0",
                   f"полный дубль {orig}: пациент, врач, время и состав услуг совпадают; "
                   f"платежа нет" if not has_pay else "дубль имеет платёж, проверить",
                   "высокая" if not has_pay else "средняя")

        # ---- цена ниже прайса ----
        elif f.error_type == "price_below_list":
            m = items.item_id == ids[0]
            r = items[m].iloc[0]
            exp = round(r.list_price_at_date * (1 - r.discount_pct / 100) / 100) * 100
            vid = r.visit_id
            pay = pays.loc[pays.visit_id == vid, "amount"].sum()
            other = items[(items.visit_id == vid) & ~m].revenue.sum()
            confirmed = abs(pay - (other + exp * r.qty)) <= TOL if pay else False
            items.loc[m, ["price_fact", "revenue"]] = [exp, exp * r.qty]
            record(f.finding_id, f.error_type, "цена по прайсу", ids[0],
                   f"цена {r.price_fact:,.0f}, выручка {r.revenue:,.0f}", f"цена {exp:,.0f}, выручка {exp * r.qty:,.0f}",
                   "платёж по визиту равен выручке при цене по прайсу" if confirmed else
                   "нет подтверждающего платежа; цена приведена к прайсу",
                   "высокая" if confirmed else "средняя")

        # ---- нет оплаты ----
        elif f.error_type == "no_payment":
            vid = ids[0]
            v = visits[visits.visit_id == vid].iloc[0]
            amt = items.loc[items.visit_id == vid, "revenue"].sum()
            pays = pd.concat([pays, pd.DataFrame([dict(
                payment_id=f"PAY{next_pay:07d}", visit_id=vid, pay_date=v.visit_date, amount=amt,
                payer_type="self", payer_name="Пациент")])], ignore_index=True)
            next_pay += 1
            record(f.finding_id, f.error_type, "восстановить платёж в ERP", vid, "платежа нет",
                   f"платёж {amt:,.0f} от {v.visit_date}",
                   "поступление этой суммы есть в 1С в день визита (ДДС сходится после восстановления)",
                   "высокая")

        # ---- выручка 1С != ERP ----
        elif f.error_type == "revenue_mismatch":
            cand = post[post.posting_id.isin(ids)]
            pref = cand[cand.debit == "62.01"]
            pool = pref if len(pref) else cand
            pool = pool[pool.amount >= abs(f.amount_rub)]
            if pool.empty:
                manual.append(dict(finding_id=f.finding_id, error_type=f.error_type, record_ids=f.record_ids,
                                   amount_rub=f.amount_rub, reason="не удалось выбрать проводку из кандидатов"))
                record(f.finding_id, f.error_type, "-", f.record_ids, "", "", "нет подходящей проводки",
                       "-", "manual_review")
                continue
            i = pool.amount.idxmax()
            old = post.loc[i, "amount"]
            post.loc[i, "amount"] = round(old - f.amount_rub, 2)
            record(f.finding_id, f.error_type, "скорректировать проводку реализации", post.loc[i, "posting_id"],
                   f"{old:,.2f}", f"{post.loc[i, 'amount']:,.2f}",
                   f"разница ERP-1С за день {f.amount_rub:+,.0f}; проводка выбрана из кандидатов "
                   f"(самоплата, наибольшая сумма) - подтвердить бухгалтером", "средняя", "applied_review")

        # ---- неверный период ----
        elif f.error_type == "wrong_period":
            i = post.index[post.posting_id == ids[0]][0]
            old = post.loc[i, "date"]
            post.loc[i, "date"] = f.fix_date
            record(f.finding_id, f.error_type, "вернуть дату реализации", ids[0], old, f.fix_date,
                   "дата визитов в ERP совпадает с исправленной; сумма проводки = сумме расхождения", "высокая")

        # ---- задвоенное поступление ----
        elif f.error_type == "duplicate_receipt":
            cand = post[post.posting_id.isin(ids)]
            same = cand[(cand.amount - f.amount_rub).abs() <= TOL]
            drop = same.posting_id.max()
            amt = post.loc[post.posting_id == drop, "amount"].iloc[0]
            post = post[post.posting_id != drop]
            record(f.finding_id, f.error_type, "удалить лишнее поступление", drop, f"{amt:,.2f}", "удалено",
                   "платежей ERP на эту сумму нет; поступление проведено дважды", "высокая")

        # ---- врач ----
        elif f.error_type == "no_doctor":
            manual.append(dict(finding_id=f.finding_id, error_type=f.error_type, record_ids=f.record_ids,
                               amount_rub=f.amount_rub,
                               reason="врача нельзя восстановить по данным; выручка остаётся без врача, "
                                      "нужен ручной разбор с клиникой"))
            record(f.finding_id, f.error_type, "-", f.record_ids, "врач не указан", "врач не указан",
                   "автоматическое исправление невозможно", "-", "manual_review")

    # ---- пересчёт 1С ----
    oper = post[post.doc_type != "Закрытие месяца"].copy()
    closing = rebuild_closing(oper)
    closing.insert(0, "posting_id", [f"PSTC{i:06d}" for i in range(1, len(closing) + 1)])
    fixed_post = pd.concat([oper, closing], ignore_index=True)
    fixed_post = fixed_post.sort_values(["date", "clinic_id", "doc_type"], kind="stable").reset_index(drop=True)
    osv = rebuild_osv(fixed_post, osv_dirty)
    chk = osv.groupby(["period", "clinic_id"])[["closing_debit", "closing_credit"]].sum()
    assert (chk.closing_debit - chk.closing_credit).abs().max() < 1, "ОСВ не сходится после исправлений"
    cash = rebuild_cash(fixed_post)

    # ---- запись ----
    for name in ["clinics", "doctors", "services", "patients", "price_history", "payroll_1c",
                 "materials_1c", "budget_revenue", "budget_opex"]:
        shutil.copy(DIRTY / f"{name}.csv", CORR / f"{name}.csv")
    enc = dict(index=False, encoding="utf-8-sig")
    visits.to_csv(CORR / "visits.csv", **enc)
    items.to_csv(CORR / "visit_items.csv", **enc)
    pays.to_csv(CORR / "visit_payments.csv", **enc)
    fixed_post.to_csv(CORR / "postings_1c.csv", **enc)
    osv.to_csv(CORR / "osv_1c.csv", **enc)
    cash.to_csv(CORR / "cashflow_1c.csv", **enc)
    pd.DataFrame(log).to_csv(OUT / "corrections_log.csv", **enc)
    pd.DataFrame(manual, columns=["finding_id", "error_type", "record_ids", "amount_rub", "reason"]
                 ).to_csv(OUT / "unresolved.csv", **enc)

    # ---- сводка ----
    lg = pd.DataFrame(log)
    print(f"Исправлений применено: {lg.status.str.startswith('applied').sum()} из {len(lg)}  "
          f"(с пометкой «подтвердить»: {(lg.status == 'applied_review').sum()}, "
          f"ручной разбор: {(lg.status == 'manual_review').sum()})")
    print(lg.groupby(["error_type", "confidence"]).size().to_string())
    onec = fixed_post[(fixed_post.doc_type == "Реализация услуг") & (fixed_post.credit == "90.01")].amount.sum()
    print(f"\nВыручка ERP: {rev_before / 1e6:,.3f} -> {items.revenue.sum() / 1e6:,.3f} млн ₽;  "
          f"1С: {onec / 1e6:,.3f} млн ₽;  разница ERP-1С: {items.revenue.sum() - onec:,.0f} ₽")
    b = osv[(osv.account == "62.01") & ((osv.closing_credit > TOL) | (osv.closing_debit > TOL))]
    print(f"Периодов с аномальным сальдо 62.01: {len(b)} (было 17)")


if __name__ == "__main__":
    main()

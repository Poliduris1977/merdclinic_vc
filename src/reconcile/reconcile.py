"""
Шаг 6: модуль сверки ERP и бухгалтерии (1С).

Читает «испорченные» данные из data/dirty/, ключ ответов НЕ использует.
Пишет в data/output/:
  findings.csv             - таблица находок (одна строка = одна проблема)
  reconciliation_report.md - отчёт для руководителя
А затем отдельным шагом сравнивает находки с ключом ответов
(data/answer_key/injected_errors.csv) и пишет оценку качества
в data/answer_key/reconcile_evaluation.md.

Логика: «объясняй и вычитай». Сначала находятся первопричины на уровне ERP
(дубли, визиты без врача, цены ниже прайса, потерянные оплаты). Их влияние
вычитается из расхождения ERP-1С, и только остаток разбирается как проблемы
самой бухгалтерии (смещение периода, ручная правка реализации, задвоенное
поступление). Так одна ошибка не порождает пять находок.

Допуск на округление: 100 руб. на ячейку сравнения.

Запуск:  python src/reconcile/reconcile.py
"""
import argparse
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
DIRTY = ROOT / "data" / "dirty"
OUT = ROOT / "data" / "output"
KEY = ROOT / "data" / "answer_key"

TOL = 100.0                     # допуск, руб.
CUTOFF = pd.Timestamp("2025-12-31")
DMS_GRACE_FROM = pd.Timestamp("2025-11-01")   # визиты с этой даты могут ждать оплату ДМС
EVAL = True
STR = {"debit": str, "credit": str, "account": str}

ACTIONS = {
    "duplicate_visit": "Удалить дубль в ERP; проверить, не создан ли дубль и в реализации 1С",
    "no_doctor": "Указать врача в визите; пересчитать мотивацию врача за период",
    "price_below_list": "Выяснить причину цены ниже прайса; оформить скидку или исправить цену",
    "no_payment": "Найти платёж по визиту (касса, эквайринг) и привязать к визиту",
    "revenue_mismatch": "Сверить документ реализации в 1С с визитами ERP; откатить ручную правку",
    "wrong_period": "Перепровести реализацию в правильный период; пересчитать закрытие месяца",
    "duplicate_receipt": "Удалить задвоенное поступление в 1С; проверить выписку банка/эквайринга",
    "missing_receipt": "Найти платёж ERP в выписке банка и провести поступление в 1С",
    "dms_receipt_mismatch": "Сверить оплату страховой с актом и платёжкой",
    "price_list_mismatch": "Проверить прайс на дату визита",
    "abnormal_balance": "Разобрать причину по связанным находкам, затем скорректировать проводки",
    "receivable_mismatch": "Сверить дебиторку ДМС по актам со страховыми",
}


def sev(amount):
    a = abs(amount)
    return "high" if a >= 100_000 else "medium" if a >= 30_000 else "low"


def load():
    v = pd.read_csv(DIRTY / "visits.csv", parse_dates=["visit_date"])
    it = pd.read_csv(DIRTY / "visit_items.csv")
    pays = pd.read_csv(DIRTY / "visit_payments.csv", parse_dates=["pay_date"])
    post = pd.read_csv(DIRTY / "postings_1c.csv", dtype=STR, parse_dates=["date"])
    osv = pd.read_csv(DIRTY / "osv_1c.csv", dtype=STR | {c: float for c in [
        "opening_debit", "opening_credit", "turnover_debit", "turnover_credit",
        "closing_debit", "closing_credit"]})
    hist = pd.read_csv(DIRTY / "price_history.csv", parse_dates=["valid_from"])
    return v, it, pays, post, osv, hist


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    visits, items, pays, post, osv, hist = load()
    F = []   # находки

    def add(rule, etype, period, clinic, direction, ids, amount, evidence, root=True, **extra):
        F.append(dict(rule=rule, error_type=etype, severity=sev(amount), period=period,
                      clinic_id=clinic, direction_code=direction, record_ids="|".join(ids),
                      amount_rub=round(float(amount), 2), evidence=evidence,
                      suggested_action=ACTIONS[etype], is_root_cause=root, **extra))

    vrev = items.groupby("visit_id").revenue.sum()
    visits["rev"] = visits.visit_id.map(vrev).fillna(0.0)
    done = visits[visits.status == "completed"].copy()
    done["period"] = done.visit_date.dt.strftime("%Y-%m")

    # ===== R1. Дубли визитов (тот же пациент, врач, время и состав услуг) =====
    sig = items.sort_values("item_id").groupby("visit_id").apply(
        lambda g: tuple(zip(g.service_id, g.qty, g.price_fact)), include_groups=False)
    done["sig"] = done.visit_id.map(sig)
    done["doc_key"] = done.doctor_id.fillna("")
    dup_keys = ["patient_id", "doc_key", "clinic_id", "visit_date", "visit_time", "sig"]
    grp = done[done.duplicated(dup_keys, keep=False)].sort_values("visit_id")
    dup_ids = set()
    for _, g in grp.groupby(dup_keys, sort=False):
        orig, *dups = g.itertuples()
        for d in dups:
            dup_ids.add(d.visit_id)
            add("R1_duplicates", "duplicate_visit", d.period, d.clinic_id, d.direction_code,
                [orig.visit_id, d.visit_id], d.rev,
                f"Визиты {orig.visit_id} и {d.visit_id}: один пациент, врач, дата {d.visit_date:%d.%m.%Y} "
                f"{d.visit_time}, одинаковые услуги; у {d.visit_id} нет платежа")

    # ===== R2. Завершённые визиты без врача =====
    for r in done[done.doctor_id.isna()].itertuples():
        add("R2_no_doctor", "no_doctor", r.period, r.clinic_id, r.direction_code, [r.visit_id], r.rev,
            f"Визит {r.visit_id} от {r.visit_date:%d.%m.%Y} на {r.rev:,.0f} ₽ без врача: "
            f"выручка не попадает в расчёт врача")

    # ===== R3. Цена ниже прайса при нулевой/меньшей скидке =====
    im = items.merge(visits[["visit_id", "visit_date", "clinic_id", "direction_code", "status"]], on="visit_id")
    im["period"] = im.visit_date.dt.strftime("%Y-%m")
    im["expected_price"] = (im.list_price_at_date * (1 - im.discount_pct / 100) / 100).round() * 100
    bad = im[im.price_fact < im.expected_price - TOL].copy()
    price_adj = {}
    for r in bad.itertuples():
        delta = (r.expected_price - r.price_fact) * r.qty
        price_adj[(r.visit_date, r.clinic_id, r.direction_code)] = \
            price_adj.get((r.visit_date, r.clinic_id, r.direction_code), 0.0) + delta
        add("R3_price_vs_list", "price_below_list", r.period, r.clinic_id, r.direction_code,
            [r.item_id], -delta,
            f"Позиция {r.item_id} (визит {r.visit_id}): цена {r.price_fact:,.0f} при прайсе "
            f"{r.list_price_at_date:,.0f} и скидке {r.discount_pct:.0f}%; недобор {delta:,.0f} ₽")

    # прайс на дату визита против истории цен (контроль справочника)
    hs = hist.sort_values("valid_from")
    m = pd.merge_asof(im.sort_values("visit_date"), hs, left_on="visit_date", right_on="valid_from",
                      by="service_id", suffixes=("", "_hist"))
    pl = m[m.list_price_at_date != m.list_price]
    for r in pl.itertuples():
        add("R3b_price_history", "price_list_mismatch", r.period, r.clinic_id, r.direction_code,
            [r.item_id], r.list_price_at_date - r.list_price,
            f"Прайс в позиции {r.list_price_at_date} != истории цен {r.list_price}")

    # ===== R4. Закрытые визиты без оплаты =====
    paid_ids = set(pays.visit_id)
    unpaid = done[~done.visit_id.isin(paid_ids) & ~done.visit_id.isin(dup_ids)]
    old_unpaid = unpaid[unpaid.visit_date < DMS_GRACE_FROM]
    nopay_adj = {}
    for r in old_unpaid.itertuples():
        nopay_adj[(r.visit_date, r.clinic_id)] = nopay_adj.get((r.visit_date, r.clinic_id), 0.0) + r.rev
        add("R4_no_payment", "no_payment", r.period, r.clinic_id, r.direction_code, [r.visit_id], r.rev,
            f"Визит {r.visit_id} от {r.visit_date:%d.%m.%Y} на {r.rev:,.0f} ₽ закрыт, платежа нет; "
            f"срок оплаты ДМС по таким визитам давно истёк")
    young_unpaid = unpaid[unpaid.visit_date >= DMS_GRACE_FROM]

    # ===== R5. Выручка ERP vs 1С по дням / клиникам / направлениям =====
    erp = im[im.status == "completed"].groupby(["visit_date", "clinic_id", "direction_code"]).revenue.sum()
    # вычитаем влияние уже найденных первопричин
    erp_adj = erp.copy()
    for d in done[done.visit_id.isin(dup_ids)].itertuples():
        k = (d.visit_date, d.clinic_id, d.direction_code)
        erp_adj[k] = erp_adj.get(k, 0.0) - d.rev
    for k, delta in price_adj.items():
        erp_adj[k] = erp_adj.get(k, 0.0) + delta
    rev1c_posts = post[(post.doc_type == "Реализация услуг") & (post.credit == "90.01")]
    c1 = rev1c_posts.groupby(["date", "clinic_id", "direction_code"]).amount.sum()
    c1.index.names = erp_adj.index.names = ["visit_date", "clinic_id", "direction_code"]
    cell = pd.concat([c1.rename("onec"), erp_adj.rename("erp")], axis=1).fillna(0.0)
    cell["diff"] = (cell.onec - cell.erp).round(2)
    small = cell[(cell["diff"].abs() > 0) & (cell["diff"].abs() <= TOL)]
    big = cell[cell["diff"].abs() > TOL].reset_index()

    used = set()
    neg = big[big["diff"] < 0]
    pos = big[big["diff"] > 0]
    posts_in = lambda d, c, dr: rev1c_posts[(rev1c_posts.date == d) & (rev1c_posts.clinic_id == c) &
                                            (rev1c_posts.direction_code == dr)]
    # --- 5a. смещение периода: минус в одной ячейке, такой же плюс в начале следующего месяца ---
    for n in neg.itertuples():
        cand = pos[(pos.clinic_id == n.clinic_id) & (pos.direction_code == n.direction_code) &
                   (pos.visit_date > n.visit_date) & (pos.visit_date <= n.visit_date + pd.Timedelta(days=6)) &
                   (pos.visit_date.dt.month != n.visit_date.month) &
                   ((pos["diff"] + n.diff).abs() <= TOL)]
        if cand.empty:
            continue
        p = cand.iloc[0]
        used.add(n.Index)
        used.add(p.name)
        ps = posts_in(p.visit_date, p.clinic_id, p.direction_code)
        exact = ps[(ps.amount - p["diff"]).abs() <= TOL]
        ids = (exact if len(exact) else ps).posting_id.tolist()
        add("R5_revenue_period", "wrong_period", n.visit_date.strftime("%Y-%m"), n.clinic_id,
            n.direction_code, ids, p["diff"],
            f"Выручка {p['diff']:,.0f} ₽ есть в ERP за {n.visit_date:%d.%m.%Y}, но проведена в 1С "
            f"{p.visit_date:%d.%m.%Y}: период {n.visit_date:%m.%Y} занижен, "
            f"{p.visit_date:%m.%Y} завышен", affects_6201=True,
            fix_date=n.visit_date.date().isoformat())
    # --- 5b. остаток: выручка в 1С не равна ERP ---
    for r in big.itertuples():
        if r.Index in used:
            continue
        ids = posts_in(r.visit_date, r.clinic_id, r.direction_code).posting_id.tolist()
        add("R5_revenue_diff", "revenue_mismatch", r.visit_date.strftime("%Y-%m"), r.clinic_id,
            r.direction_code, ids, r.diff,
            f"{r.visit_date:%d.%m.%Y}: выручка 1С {r.onec:,.0f} ₽, ERP (с учётом найденных ошибок) "
            f"{r.erp:,.0f} ₽, разница {r.diff:+,.0f} ₽. Проводки-кандидаты: {', '.join(ids)}",
            affects_6201=True)

    # ===== R6. Поступления: ERP-платежи vs проводки 1С =====
    cash = post[post.debit == "51"]
    e_self = pays[pays.payer_type == "self"].merge(visits[["visit_id", "clinic_id"]], on="visit_id")
    e_self = e_self.groupby(["pay_date", "clinic_id"]).amount.sum()
    c_self = cash[cash.doc_type == "Поступление от пациентов (эквайринг)"].groupby(["date", "clinic_id"]).amount.sum()
    e_self.index.names = c_self.index.names = ["d", "clinic_id"]
    adj = e_self.copy()
    for k, v in nopay_adj.items():
        adj[k] = adj.get(k, 0.0) + v
    cc = pd.concat([c_self.rename("onec"), adj.rename("erp")], axis=1).fillna(0.0)
    cc["diff"] = (cc.onec - cc.erp).round(2)
    for (d, c), r in cc[cc["diff"].abs() > TOL].iterrows():
        ids = cash[(cash.date == d) & (cash.clinic_id == c) &
                   (cash.doc_type == "Поступление от пациентов (эквайринг)")].posting_id.tolist()
        if r["diff"] > 0:
            add("R6_cash_self", "duplicate_receipt", d.strftime("%Y-%m"), c, "", ids, r["diff"],
                f"{d:%d.%m.%Y}: поступлений в 1С {r.onec:,.0f} ₽, платежей в ERP {r.erp:,.0f} ₽; "
                f"лишние {r['diff']:,.0f} ₽ (проводки {', '.join(ids)})", affects_6201=True)
        else:
            add("R6_cash_self", "missing_receipt", d.strftime("%Y-%m"), c, "", ids, r["diff"],
                f"{d:%d.%m.%Y}: платежи ERP на {-r['diff']:,.0f} ₽ не проведены в 1С")

    e_dms = pays[pays.payer_type == "dms"].merge(visits[["visit_id", "clinic_id"]], on="visit_id")
    e_dms = e_dms.groupby(["pay_date", "clinic_id", "payer_name"]).amount.sum()
    c_dms = cash[(cash.doc_type == "Поступление от страховой") &
                 (cash.comment.fillna("") != "Оплата акта за декабрь 2024")]
    c_dms = c_dms.groupby(["date", "clinic_id", "counterparty"]).amount.sum()
    e_dms.index.names = c_dms.index.names = ["d", "clinic_id", "payer"]
    dd = pd.concat([c_dms.rename("onec"), e_dms.rename("erp")], axis=1).fillna(0.0)
    dd["diff"] = (dd.onec - dd.erp).round(2)
    for (d, c, p), r in dd[dd["diff"].abs() > TOL].iterrows():
        add("R6_cash_dms", "dms_receipt_mismatch", d.strftime("%Y-%m"), c, "", [], r["diff"],
            f"{d:%d.%m.%Y} {p}: 1С {r.onec:,.0f} ₽, ERP {r.erp:,.0f} ₽")

    # ===== R7. Дебиторка ДМС: ОСВ 62.02 на 31.12 vs неоплаченные визиты ERP =====
    dms_1c = osv[(osv.period == "2025-12") & (osv.account == "62.02")].set_index("clinic_id").closing_debit
    dms_erp = young_unpaid.groupby("clinic_id").rev.sum()
    rec = pd.concat([dms_1c.rename("onec"), dms_erp.rename("erp")], axis=1).fillna(0.0)
    rec["diff"] = rec.onec - rec.erp
    for c, r in rec.iterrows():
        if abs(r["diff"]) > 5 * TOL:
            add("R7_dms_receivable", "receivable_mismatch", "2025-12", c, "", [], r["diff"],
                f"62.02 в ОСВ {r.onec:,.0f} ₽ против неоплаченных визитов ERP {r.erp:,.0f} ₽", root=False)

    # ===== R8. Аномальное сальдо 62.01 (симптом) =====
    roots = pd.DataFrame(F)
    b = osv[(osv.account == "62.01") & ((osv.closing_credit > TOL) | (osv.closing_debit > TOL))]
    for c, g in b.groupby("clinic_id"):
        g = g.sort_values("period")
        first, last = g.iloc[0], g.iloc[-1]
        linked = []
        if len(roots) and "affects_6201" in roots:
            linked = roots[(roots.clinic_id == c) & (roots.affects_6201 == True)].index.tolist()  # noqa: E712
        amount = last.closing_debit - last.closing_credit
        add("R8_osv_balance", "abnormal_balance", first.period, c, "", [], amount,
            f"Счёт 62.01 (расчёты с пациентами) должен закрываться в ноль. Отклонение с {first.period}, "
            f"на конец периода {'дебетовое' if amount > 0 else 'кредитовое'} сальдо {abs(amount):,.0f} ₽. "
            f"Связанные находки: @@LINKED@@", root=False,
            _linked=linked)

    # ===== Сборка =====
    findings = pd.DataFrame(F)
    findings["_orig"] = findings.index
    if "affects_6201" in findings:
        findings = findings.drop(columns="affects_6201")
    order = {"high": 0, "medium": 1, "low": 2}
    findings["_o"] = findings.severity.map(order)
    findings["_abs"] = -findings.amount_rub.abs()
    findings = findings.sort_values(["is_root_cause", "_o", "_abs"],
                                    ascending=[False, True, True]).reset_index(drop=True)
    findings.insert(0, "finding_id", [f"F{i + 1:03d}" for i in range(len(findings))])
    orig2id = dict(zip(findings._orig, findings.finding_id))
    for i, r in findings.iterrows():
        if "@@LINKED@@" in r.evidence:
            ids = ", ".join(orig2id[j] for j in r._linked) if isinstance(r._linked, list) and r._linked else "не найдены"
            findings.at[i, "evidence"] = r.evidence.replace("@@LINKED@@", ids)
    findings = findings.drop(columns=["_o", "_abs", "_orig", "_linked"], errors="ignore")
    findings.to_csv(OUT / "findings.csv", index=False, encoding="utf-8-sig")

    write_report(findings, small, young_unpaid, rec, len(done), len(items))
    if EVAL:
        evaluate(findings)


def write_report(f, small, young_unpaid, rec, n_visits, n_items):
    roots = f[f.is_root_cause]
    L = ["# Отчёт сверки ERP и 1С за 2025 год", ""]
    L.append(f"Проверено: {n_visits:,} завершённых визитов, {n_items:,} услуг, проводки 1С, платежи, ОСВ.")
    L.append(f"**Найдено первопричин: {len(roots)}**, симптомов в ОСВ: {len(f) - len(roots)}.")
    L.append(f"Сумма под риском (по модулю): **{roots.amount_rub.abs().sum():,.0f} ₽**.")
    L.append("")
    L += ["## Сводка по типам", "", "| Тип | Кол-во | Сумма, ₽ |", "|---|---|---|"]
    names = {"duplicate_visit": "Дубли визитов", "no_doctor": "Визит без врача",
             "price_below_list": "Цена ниже прайса", "no_payment": "Визит без оплаты",
             "revenue_mismatch": "Выручка 1С ≠ ERP", "wrong_period": "Неверный период",
             "duplicate_receipt": "Задвоенное поступление", "missing_receipt": "Платёж не проведён",
             "dms_receipt_mismatch": "Оплата ДМС не сходится", "price_list_mismatch": "Прайс не по истории"}
    for t, g in roots.groupby("error_type"):
        L.append(f"| {names.get(t, t)} | {len(g)} | {g.amount_rub.abs().sum():,.0f} |")
    L += ["", "## Находки", "", "| ID | Важность | Период | Клиника | Тип | Сумма, ₽ | Что и где |",
          "|---|---|---|---|---|---|---|"]
    for r in roots.itertuples():
        L.append(f"| {r.finding_id} | {r.severity} | {r.period} | {r.clinic_id} | "
                 f"{names.get(r.error_type, r.error_type)} | {r.amount_rub:,.0f} | {r.evidence} |")
    L += ["", "## Что делать", ""]
    for t, g in roots.groupby("error_type"):
        L.append(f"- **{names.get(t, t)}** ({len(g)}): {g.suggested_action.iloc[0]}.")
    sym = f[~f.is_root_cause]
    if len(sym):
        L += ["", "## Симптомы в ОСВ", ""]
        for r in sym.itertuples():
            L.append(f"- {r.finding_id}, {r.clinic_id}: {r.evidence}")
    L += ["", "## Что проверено и принято как норма", "",
          f"- Расхождения до {TOL:.0f} ₽ на ячейку (округления): {len(small)} шт., "
          f"сумма по модулю {small['diff'].abs().sum():,.0f} ₽. Не ошибки.",
          f"- Неоплаченные визиты с 01.11.2025 ({len(young_unpaid)} шт., "
          f"{young_unpaid.rev.sum():,.0f} ₽): ДМС платит через 30-45 дней, срок не вышел.",
          "- Дебиторка ДМС в ОСВ (62.02) на 31.12 сходится с неоплаченными визитами ERP: "
          + ", ".join(f"{c}: {r.onec:,.0f} против {r.erp:,.0f}" for c, r in rec.iterrows()) + " ₽.",
          "- Оплаты ДМС за декабрь 2024 (акт прошлого года) исключены из сверки.",
          "", "## Ограничения", "",
          "- Потерянную самооплату в ноябре-декабре нельзя отличить от ожидаемой оплаты ДМС: "
          "в ERP нет признака плательщика у неоплаченного визита. Для таких визитов нужен признак плательщика.",
          "- Сверка делается на уровне день / клиника / направление; внутри ячейки проводки-кандидаты "
          "перечислены, но точную проводку подтверждает бухгалтер.",
          "- Не проверялись: НДС, налог на прибыль, расчёт зарплаты по врачам."]
    import re
    text = re.sub(r"(?<=\d),(?=\d{3})", " ", "\n".join(L))   # разряды пробелом
    (OUT / "reconciliation_report.md").write_text(text, encoding="utf-8")


def evaluate(f):
    """Оценка по ключу. Строгое совпадение: тот же тип и пересечение id записей (если у ошибки в ключе нет id,
    например потерянное поступление, то тот же тип, клиника и период). Дополнительно: пропущенные ошибки,
    у которых есть находка-симптом (другого типа) в том же периоде, и находки без пары в ключе, лежащие
    в периоде пропущенной ошибки (вероятные симптомы, а не ложные срабатывания)."""
    kp = KEY / "injected_errors.csv"
    if not kp.exists():
        return
    key = pd.read_csv(kp)
    errs = key[key.error_id != "BENIGN"].copy()
    benign = (key.error_id == "BENIGN").sum()
    roots = f[f.is_root_cause].copy()
    roots["key_type"] = roots.error_type.replace({"duplicate_receipt": "negative_balance"})
    tp, matched_f, miss = [], set(), []
    for e in errs.itertuples():
        ids = {x for x in str(e.record_ids).split("|") if x and x != "nan"}
        same = roots[(roots.key_type == e.error_type) & (roots.clinic_id == e.clinic_id) &
                     (roots.period == e.period)]
        by_id = roots[(roots.key_type == e.error_type) &
                      roots.record_ids.map(lambda s: bool(ids & set(str(s).split("|"))))]
        # id совпали, либо у ключа/находки нет id (потерянное поступление, расхождение по ДМС) и совпали тип, клиника, период
        no_id = same[same.record_ids.fillna("").eq("") | (not ids)]
        hit = pd.concat([by_id, no_id]).drop_duplicates("finding_id")
        if len(hit):
            tp.append(e.error_id)
            matched_f |= set(hit.finding_id)
        else:
            miss.append(e.error_id)
    fp = roots[~roots.finding_id.isin(matched_f)]
    prec = len(matched_f) / max(len(roots), 1)
    rec = len(tp) / len(errs)

    miss_rows = errs[errs.error_id.isin(miss)]
    fp_ids_used = set()
    symptom = {}
    for e in miss_rows.itertuples():
        if pd.isna(e.amount_effect_rub):
            continue
        near = fp[(fp.period == e.period) &
                  (fp.amount_rub.abs().sub(abs(e.amount_effect_rub)).abs() <= 0.05 * abs(e.amount_effect_rub))]
        if len(near):
            symptom[e.error_id] = sorted(set(near.error_type))
            fp_ids_used |= set(near.finding_id)
    fp_symptom = fp[fp.finding_id.isin(fp_ids_used)]
    fp_real = fp[~fp.finding_id.isin(fp_ids_used)]

    L = ["# Оценка сверки по ключу ответов", "",
         f"- Внесено ошибок: {len(errs)}, безобидных округлений: {benign}",
         f"- Найдено (тип и запись совпали): {len(tp)} из {len(errs)} (полнота {rec:.0%})",
         f"- Находок-первопричин: {len(roots)}, из них подтверждённых ключом: {len(matched_f)} "
         f"(точность {prec:.0%})",
         f"- Ложных срабатываний: {len(fp_real)}" + (f" (ещё {len(fp_symptom)} находок без пары в ключе: вероятные симптомы пропущенных ошибок)" if len(fp_symptom) else ""),
         f"- Пропущено: {', '.join(miss) or 'нет'}", ""]
    if miss:
        L.append("## Не распознанные по типу: заметно ли расхождение")
        for e in miss_rows.itertuples():
            sy = symptom.get(e.error_id)
            L.append(f"- {e.error_id} {e.error_type} ({e.clinic_id}, {e.period}): "
                     + (f"расхождение замечено находками типов {', '.join(sy)}, причина не определена" if sy
                        else "не замечена вообще"))
        L.append("")
    if len(fp):
        L.append("## Находки без пары в ключе")
        for r in fp_symptom.itertuples():
            L.append(f"- (вероятный симптом пропущенной ошибки) {r.finding_id} {r.error_type} {r.period} {r.clinic_id}: {r.evidence}")
        for r in fp_real.itertuples():
            L.append(f"- (ложное срабатывание) {r.finding_id} {r.error_type} {r.period} {r.clinic_id}: {r.evidence}")
    (KEY / "reconcile_evaluation.md").write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Сверка ERP и 1С")
    ap.add_argument("--input", default="dirty", help="папка в data/ с данными (dirty | corrected)")
    ap.add_argument("--out", default="output", help="папка в data/ для результатов")
    ap.add_argument("--key", default="answer_key", help="папка в data/ с ключом ответов (для оценки)")
    args = ap.parse_args()
    DIRTY = ROOT / "data" / args.input
    OUT = ROOT / "data" / args.out
    KEY = ROOT / "data" / args.key
    EVAL = args.input != "corrected"
    main()

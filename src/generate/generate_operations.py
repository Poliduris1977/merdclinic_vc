"""
Шаг 2 генератора: операционный слой ERP за календарный 2025 год.

Читает справочники из data/raw/ (их создаёт generate_reference.py) и пишет:
  patients.csv       - пациенты
  visits.csv         - визиты (completed / cancelled / no_show)
  visit_items.csv    - оказанные услуги в завершённых визитах
  price_history.csv  - история изменения прайса (индексации по направлениям)

Данные ЧИСТЫЕ: ошибки закладываются отдельным шагом (inject_errors.py),
чтобы потом можно было сравнить «эталон» и «испорченную» версию и
посчитать, сколько ошибок нашла сверка.

Что моделируется:
  - сезонность по месяцам и направлениям, рост спроса в течение года
  - рабочие дни Пн-Сб, праздники, отпуска врачей, дата найма
  - популярность врача зависит от грейда
  - воронка пациента: первичный приём (консультация) -> повторные визиты
  - индексации прайса по направлениям в разные месяцы
  - скидки (чаще в стоматологии, косметологии и семейной медицине)

Запуск:  python src/generate/generate_operations.py
"""
import random
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 42
YEAR = 2025
RAW = Path(__file__).resolve().parents[2] / "data" / "raw"

rng = random.Random(SEED)
nprng = np.random.default_rng(SEED)

# --- Сезонность ------------------------------------------------------------
MONTH_FACTOR = {1: 0.62, 2: 0.95, 3: 1.00, 4: 1.00, 5: 0.82, 6: 0.95,
                7: 0.88, 8: 0.85, 9: 1.08, 10: 1.18, 11: 1.18, 12: 1.05}
DIR_MONTH_ADJ = {
    "COSM":  {6: 0.9, 7: 0.85, 11: 1.10, 12: 1.20},
    "LASER": {6: 0.85, 7: 0.80, 8: 0.80, 10: 1.10, 11: 1.15, 12: 1.10},
    "PLAS":  {5: 0.90, 9: 1.15, 10: 1.15},
}
TREND_PER_MONTH = 0.008  # рост спроса +0.8% в месяц

# Базовое число визитов на врача в рабочий день
BASE_VISITS = {"DENT": 5.2, "COSM": 5.8, "PLAS": 1.7, "FAM": 12.0, "LASER": 4.4}
# Вероятность повторного визита (пациент уже есть в базе клиники/направления)
P_REPEAT = {"DENT": 0.62, "COSM": 0.55, "FAM": 0.65, "PLAS": 0.15, "LASER": 0.12}
# Размер «стартовой базы» пациентов на врача (пациенты, пришедшие до 2025 года)
SEED_POOL_PER_DOCTOR = 220
P_CROSS_SELL = 0.08

# Скидки: (вероятность скидки, [значения], [веса])
DISCOUNT = {
    "DENT": (0.24, [5, 10, 15], [5, 3, 1]),
    "COSM": (0.26, [5, 10, 15], [5, 3, 1]),
    "FAM":  (0.20, [5, 10], [4, 1]),
    "PLAS": (0.08, [5, 10], [3, 1]),
    "LASER": (0.10, [5, 10], [3, 1]),
}

# Индексации прайса: направление -> (дата, коэффициент)
PRICE_CHANGES = {
    "FAM":   (date(YEAR, 2, 1), 1.07),
    "DENT":  (date(YEAR, 4, 1), 1.05),
    "COSM":  (date(YEAR, 4, 1), 1.04),
    "LASER": (date(YEAR, 9, 1), 1.03),
}

# Веса услуг при выборе для повторного визита (по названию)
REPEAT_WEIGHTS = {
    "Профессиональная гигиена": 3, "Лечение кариеса (пломба)": 5,
    "Лечение пульпита (1 канал)": 2.5, "Удаление зуба простое": 1.2,
    "Имплантация (имплант + установка)": 0.6, "Коронка металлокерамика": 1.0,
    "Коронка цирконий": 0.6, "Виниры (1 единица)": 0.3, "Отбеливание ZOOM": 0.5,
    "Брекеты (установка, 1 челюсть)": 0.4,
    "Ботулинотерапия (зона)": 4, "Контурная пластика (1 мл)": 2.5,
    "Биоревитализация": 2, "Химический пилинг": 2.5,
    "Аппаратная чистка лица": 3, "Лазерное омоложение": 1.5,
    "Блефаропластика": 1, "Ринопластика": 1, "Маммопластика (эндопротез)": 1,
    "Абдоминопластика": 0.6, "Липосакция (1 зона)": 1.2,
    "Приём терапевта": 6, "Приём педиатра": 5, "Приём гинеколога": 4,
    "Приём невролога": 3, "УЗИ органов брюшной полости": 2,
    "Комплексный анализ крови": 3, "Check-up программа": 0.4, "Вакцинация": 1.5,
    "Лазерная коррекция FemtoLASIK (2 глаза)": 1.8,
    "Лазерная коррекция ReLEx SMILE (2 глаза)": 1.4,
    "Послеоперационный осмотр": 3,
}
CONSULT = {  # услуга первичного приёма для новых пациентов
    "DENT": "Консультация стоматолога", "COSM": "Консультация косметолога",
    "PLAS": "Консультация пластического хирурга",
    "LASER": "Диагностика зрения перед коррекцией",
}
# Какие услуги может оказывать врач данной специальности
SPECIALTY_SERVICES = {
    "Стоматолог-терапевт": ["Профессиональная гигиена", "Лечение кариеса (пломба)",
                            "Лечение пульпита (1 канал)", "Отбеливание ZOOM"],
    "Стоматолог-ортопед": ["Коронка металлокерамика", "Коронка цирконий", "Виниры (1 единица)"],
    "Стоматолог-хирург": ["Удаление зуба простое", "Имплантация (имплант + установка)"],
    "Ортодонт": ["Брекеты (установка, 1 челюсть)", "Профессиональная гигиена"],
    "Терапевт": ["Приём терапевта", "Комплексный анализ крови", "Check-up программа",
                 "Вакцинация", "УЗИ органов брюшной полости"],
    "Педиатр": ["Приём педиатра", "Вакцинация", "Комплексный анализ крови"],
    "Гинеколог": ["Приём гинеколога", "УЗИ органов брюшной полости"],
    "Невролог": ["Приём невролога"],
}
ITEMS_PER_VISIT = {  # распределение числа услуг в повторном визите
    "DENT": [0.62, 0.28, 0.10], "COSM": [0.75, 0.20, 0.05], "FAM": [0.75, 0.20, 0.05],
    "PLAS": [1.0, 0, 0], "LASER": [1.0, 0, 0],
}
QTY_CHOICES = {
    "Лечение кариеса (пломба)": [1, 1, 2, 3],
    "Лечение пульпита (1 канал)": [1, 1, 2],
    "Удаление зуба простое": [1, 1, 2],
}
SOURCES = [("Яндекс", 28), ("Сарафанное радио", 24), ("2ГИС", 14), ("Таргет", 12),
           ("ДМС", 12), ("Повторное обращение", 4), ("Прочее", 6)]
STATUS = [("completed", 0.90), ("cancelled", 0.07), ("no_show", 0.03)]

HOLIDAYS = ({date(YEAR, 1, d) for d in range(1, 9)} | {date(YEAR, 3, 8)} |
            {date(YEAR, 5, d) for d in (1, 2, 3, 4, 8, 9, 10, 11)} |
            {date(YEAR, 6, d) for d in (12, 13, 14, 15)} |
            {date(YEAR, 11, d) for d in (1, 2, 3, 4)} | {date(YEAR, 12, 31)})


def round100(x: float) -> int:
    return int(round(x / 100.0) * 100)


def build_price_history(services: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for s in services.itertuples():
        rows.append(dict(service_id=s.service_id, valid_from=date(YEAR, 1, 1).isoformat(),
                         list_price=int(s.list_price)))
        if s.direction_code in PRICE_CHANGES:
            d, k = PRICE_CHANGES[s.direction_code]
            rows.append(dict(service_id=s.service_id, valid_from=d.isoformat(),
                             list_price=round100(s.list_price * k)))
    return pd.DataFrame(rows)


def make_price_lookup(history: pd.DataFrame):
    by_service = {}
    for sid, g in history.groupby("service_id"):
        g = g.sort_values("valid_from")
        by_service[sid] = [(date.fromisoformat(v), p) for v, p in zip(g.valid_from, g.list_price)]

    def price_at(sid: str, d: date) -> int:
        price = by_service[sid][0][1]
        for vf, p in by_service[sid]:
            if d >= vf:
                price = p
        return price
    return price_at


def pick_services(direction: str, specialty: str, svc_by_name: dict, is_new: bool) -> list:
    """Возвращает список (service_id, name) для визита."""
    allowed = SPECIALTY_SERVICES.get(specialty)
    names = [n for n in REPEAT_WEIGHTS
             if svc_by_name[n]["direction_code"] == direction and (allowed is None or n in allowed)]
    weights = [REPEAT_WEIGHTS[n] for n in names]

    chosen = []
    if is_new and direction in CONSULT:
        chosen.append(CONSULT[direction])
        if direction in ("DENT", "COSM") and rng.random() < 0.35:
            chosen.append(rng.choices(names, weights)[0])
    else:
        n_items = rng.choices([1, 2, 3], ITEMS_PER_VISIT[direction])[0]
        tries = 0
        while len(chosen) < n_items and tries < 10:
            tries += 1
            pick = rng.choices(names, weights)[0]
            if pick not in chosen:
                chosen.append(pick)
    return [(svc_by_name[n]["service_id"], n) for n in chosen]


def main():
    clinics = pd.read_csv(RAW / "clinics.csv")
    services = pd.read_csv(RAW / "services.csv")
    doctors = pd.read_csv(RAW / "doctors.csv")
    svc_by_name = {r.service_name: r._asdict() for r in services.itertuples()}

    history = build_price_history(services)
    price_at = make_price_lookup(history)

    # --- Пациенты: стартовые пулы (пришли до 2025 года) ---------------------
    patients = {}          # patient_id -> dict
    pools, pool_sets = {}, {}   # (clinic, direction) -> list / set
    pid_counter = 0

    def new_patient(first_date: date, clinic_id: str) -> str:
        nonlocal pid_counter
        pid_counter += 1
        pid = f"P{pid_counter:06d}"
        patients[pid] = dict(
            patient_id=pid, first_visit_date=first_date.isoformat(), home_clinic_id=clinic_id,
            birth_year=rng.randint(1950, 2012),
            source=rng.choices([s for s, _ in SOURCES], [w for _, w in SOURCES])[0],
        )
        return pid

    def add_to_pool(key, pid):
        if pid not in pool_sets.setdefault(key, set()):
            pool_sets[key].add(pid)
            pools.setdefault(key, []).append(pid)

    for (clinic_id, direction), g in doctors.groupby(["clinic_id", "direction_code"]):
        for _ in range(len(g) * SEED_POOL_PER_DOCTOR):
            d = date(2021, 1, 1) + timedelta(days=rng.randint(0, 1460))
            add_to_pool((clinic_id, direction), new_patient(d, clinic_id))

    # --- Врачи: отпуска и график -------------------------------------------
    doc_rows = []
    for d in doctors.itertuples():
        hire = date.fromisoformat(d.hire_date)
        start = max(hire, date(YEAR, 1, 1))
        v_start = date(YEAR, 1, 1) + timedelta(days=rng.randint(30, 300))
        doc_rows.append(dict(
            doc=d, start=start, vac=(v_start, v_start + timedelta(days=27)),
            level=BASE_VISITS[d.direction_code] * (d.popularity_weight ** 0.5),
        ))

    visits, items = [], []
    v_counter = i_counter = 0
    day = date(YEAR, 1, 1)
    while day.year == YEAR:
        if day.weekday() != 6:  # воскресенье - выходной
            wd = 1.0 if day.weekday() < 5 else 0.7
            hol = 0.3 if day in HOLIDAYS else 1.0
            trend = 1 + TREND_PER_MONTH * (day.month - 1)
            for r in doc_rows:
                d = r["doc"]
                if day < r["start"] or r["vac"][0] <= day <= r["vac"][1]:
                    continue
                season = MONTH_FACTOR[day.month] * DIR_MONTH_ADJ.get(d.direction_code, {}).get(day.month, 1.0)
                n = nprng.poisson(r["level"] * season * trend * wd * hol)
                key = (d.clinic_id, d.direction_code)
                for _ in range(n):
                    # пациент
                    is_new = True
                    pool = pools.get(key, [])
                    if pool and rng.random() < P_REPEAT[d.direction_code]:
                        pid, is_new = rng.choice(pool), False
                    else:
                        others = [k for k in pools if k[0] == d.clinic_id and k != key and pools[k]]
                        if others and rng.random() < P_CROSS_SELL:
                            pid = rng.choice(pools[rng.choice(others)])
                        else:
                            pid = new_patient(day, d.clinic_id)
                        add_to_pool(key, pid)

                    v_counter += 1
                    vid = f"V{v_counter:07d}"
                    status = rng.choices([s for s, _ in STATUS], [w for _, w in STATUS])[0]
                    visits.append(dict(
                        visit_id=vid, patient_id=pid, doctor_id=d.doctor_id, clinic_id=d.clinic_id,
                        direction_code=d.direction_code, visit_date=day.isoformat(),
                        visit_time=f"{rng.randint(9, 19):02d}:{rng.choice([0, 15, 30, 45]):02d}",
                        status=status, is_first_in_direction=is_new,
                    ))
                    if status != "completed":
                        continue
                    for sid, name in pick_services(d.direction_code, d.specialty, svc_by_name, is_new):
                        qty = rng.choice(QTY_CHOICES.get(name, [1]))
                        p_disc, vals, wts = DISCOUNT[d.direction_code]
                        disc = rng.choices(vals, wts)[0] if rng.random() < p_disc else 0
                        lp = price_at(sid, day)
                        pf = round100(lp * (1 - disc / 100))
                        i_counter += 1
                        items.append(dict(
                            item_id=f"I{i_counter:07d}", visit_id=vid, service_id=sid, qty=qty,
                            list_price_at_date=lp, discount_pct=disc, price_fact=pf, revenue=pf * qty,
                        ))
        day += timedelta(days=1)

    visits_df, items_df = pd.DataFrame(visits), pd.DataFrame(items)
    used = set(visits_df.patient_id)
    patients_df = pd.DataFrame([p for pid, p in patients.items() if pid in used])

    RAW.mkdir(parents=True, exist_ok=True)
    for name, df in [("visits", visits_df), ("visit_items", items_df),
                     ("patients", patients_df), ("price_history", history)]:
        df.to_csv(RAW / f"{name}.csv", index=False, encoding="utf-8-sig")

    # --- Контрольная сводка -------------------------------------------------
    done = visits_df[visits_df.status == "completed"]
    m = items_df.merge(visits_df[["visit_id", "visit_date", "clinic_id", "direction_code"]], on="visit_id")
    m["month"] = m.visit_date.str[:7]
    print(f"patients: {len(patients_df)}  visits: {len(visits_df)} (completed {len(done)})  items: {len(items_df)}")
    print(f"Выручка за {YEAR}: {m.revenue.sum() / 1e6:,.1f} млн ₽")
    print("\nВыручка по месяцам, млн ₽:")
    print((m.groupby("month").revenue.sum() / 1e6).round(1).to_string())
    print("\nВыручка по направлениям, млн ₽:")
    print((m.groupby("direction_code").revenue.sum() / 1e6).round(1).to_string())
    print("\nВыручка по клиникам, млн ₽:")
    print((m.groupby("clinic_id").revenue.sum() / 1e6).round(1).to_string())
    print(f"\nДоля повторных визитов: {1 - done.is_first_in_direction.mean():.1%}")
    print(f"Средняя скидка (по позициям): {items_df.discount_pct.mean():.1f}%")


if __name__ == "__main__":
    main()

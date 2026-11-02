"""
Шаг 1 генератора синтетических данных: справочники.

Создаёт в data/raw/:
  clinics.csv   - клиники сети
  services.csv  - прайс-лист и плановая себестоимость услуг
  doctors.csv   - врачи, грейды, модель оплаты

Все данные синтетические. Названия и цены вымышлены, но экономика
(доля материалов, лаборатории, процент врача) соответствует типичным
значениям для частной медицины.

Запуск:  python src/generate/generate_reference.py
"""
import random
from pathlib import Path

import pandas as pd

SEED = 42
OUT = Path(__file__).resolve().parents[2] / "data" / "raw"

# --- Клиники ---------------------------------------------------------------
# directions: какие направления представлены в клинике
CLINICS = [
    dict(clinic_id="C01", clinic_name="Клиника Невский", district="Центральный",
         directions="DENT|COSM|PLAS|FAM|LASER", rent_month=1_450_000),
    dict(clinic_id="C02", clinic_name="Клиника Север", district="Выборгский",
         directions="DENT|COSM|FAM", rent_month=780_000),
    dict(clinic_id="C03", clinic_name="Клиника Юг", district="Московский",
         directions="DENT|FAM|LASER", rent_month=820_000),
]

DIRECTIONS = {
    "DENT": "Стоматология",
    "COSM": "Косметология",
    "PLAS": "Пластическая хирургия",
    "FAM": "Семейная медицина",
    "LASER": "Лазерная коррекция зрения",
}

# --- Услуги ----------------------------------------------------------------
# (direction, service_name, price, material_cost, external_cost, duration_min)
# material_cost - расходные материалы; external_cost - лаборатория / внешние
# расходы (зуботехническая лаборатория, импланты, анализы на стороне).
SERVICES = [
    # Стоматология
    ("DENT", "Консультация стоматолога", 1_500, 20, 0, 30),
    ("DENT", "Профессиональная гигиена", 7_500, 650, 0, 60),
    ("DENT", "Лечение кариеса (пломба)", 9_000, 1_100, 0, 60),
    ("DENT", "Лечение пульпита (1 канал)", 14_000, 2_300, 0, 90),
    ("DENT", "Удаление зуба простое", 4_500, 400, 0, 30),
    ("DENT", "Имплантация (имплант + установка)", 65_000, 3_500, 26_000, 90),
    ("DENT", "Коронка металлокерамика", 24_000, 800, 7_500, 60),
    ("DENT", "Коронка цирконий", 38_000, 900, 13_000, 60),
    ("DENT", "Виниры (1 единица)", 45_000, 700, 15_000, 90),
    ("DENT", "Отбеливание ZOOM", 28_000, 3_200, 0, 90),
    ("DENT", "Брекеты (установка, 1 челюсть)", 55_000, 6_000, 12_000, 90),
    # Косметология
    ("COSM", "Консультация косметолога", 2_000, 0, 0, 30),
    ("COSM", "Ботулинотерапия (зона)", 14_000, 6_500, 0, 30),
    ("COSM", "Контурная пластика (1 мл)", 22_000, 11_000, 0, 45),
    ("COSM", "Биоревитализация", 18_000, 7_800, 0, 45),
    ("COSM", "Химический пилинг", 9_000, 1_400, 0, 45),
    ("COSM", "Аппаратная чистка лица", 8_500, 700, 0, 60),
    ("COSM", "Лазерное омоложение", 26_000, 2_200, 0, 60),
    # Пластическая хирургия
    ("PLAS", "Консультация пластического хирурга", 3_000, 0, 0, 40),
    ("PLAS", "Блефаропластика", 160_000, 12_000, 28_000, 120),
    ("PLAS", "Ринопластика", 320_000, 25_000, 55_000, 210),
    ("PLAS", "Маммопластика (эндопротез)", 420_000, 30_000, 140_000, 180),
    ("PLAS", "Абдоминопластика", 380_000, 28_000, 60_000, 240),
    ("PLAS", "Липосакция (1 зона)", 150_000, 9_000, 20_000, 120),
    # Семейная медицина
    ("FAM", "Приём терапевта", 3_500, 80, 0, 30),
    ("FAM", "Приём педиатра", 3_200, 80, 0, 30),
    ("FAM", "Приём гинеколога", 4_000, 150, 0, 40),
    ("FAM", "Приём невролога", 4_200, 60, 0, 40),
    ("FAM", "УЗИ органов брюшной полости", 3_800, 120, 0, 30),
    ("FAM", "Комплексный анализ крови", 2_900, 450, 1_200, 15),
    ("FAM", "Check-up программа", 24_000, 1_800, 9_500, 180),
    ("FAM", "Вакцинация", 2_800, 100, 1_400, 15),
    # Лазерная коррекция зрения
    ("LASER", "Диагностика зрения перед коррекцией", 4_500, 150, 0, 60),
    ("LASER", "Лазерная коррекция FemtoLASIK (2 глаза)", 145_000, 9_500, 12_000, 60),
    ("LASER", "Лазерная коррекция ReLEx SMILE (2 глаза)", 175_000, 11_000, 14_000, 60),
    ("LASER", "Послеоперационный осмотр", 2_500, 40, 0, 20),
]

# --- Врачи -----------------------------------------------------------------
# Кол-во врачей по направлениям и клиникам
DOCTORS_PLAN = {
    "DENT":  {"C01": 4, "C02": 3, "C03": 3},
    "COSM":  {"C01": 5, "C02": 4},
    "PLAS":  {"C01": 6},
    "FAM":   {"C01": 5, "C02": 3, "C03": 4},
    "LASER": {"C01": 3, "C03": 3},
}
SPECIALTY_NAMES = {
    "DENT": ["Стоматолог-терапевт", "Стоматолог-ортопед", "Стоматолог-хирург", "Ортодонт"],
    "COSM": ["Врач-косметолог", "Дерматовенеролог-косметолог"],
    "PLAS": ["Пластический хирург"],
    "FAM": ["Терапевт", "Педиатр", "Гинеколог", "Невролог"],
    "LASER": ["Офтальмолог-хирург"],
}
# Грейд -> (доля врача от выручки по умолчанию, вес популярности врача)
GRADES = {
    "junior": (0.28, 0.6),
    "middle": (0.32, 1.0),
    "senior": (0.36, 1.5),
    "lead":   (0.40, 2.0),
}
GRADE_PROBS = [("junior", 0.20), ("middle", 0.40), ("senior", 0.28), ("lead", 0.12)]

SURNAMES = ["Иванов", "Петров", "Смирнов", "Кузнецов", "Попов", "Соколов", "Лебедев",
            "Козлов", "Новиков", "Морозов", "Волков", "Алексеев", "Орлов", "Зайцев",
            "Белов", "Тихонов", "Макаров", "Фролов", "Громов", "Дьяконов", "Ширяев",
            "Крылов", "Савельев", "Нестеров", "Ковалёв", "Жуков", "Борисов", "Анисимов"]
INITIALS = ["А.", "Е.", "И.", "М.", "О.", "С.", "Т.", "Н.", "Д.", "В."]


def build_clinics() -> pd.DataFrame:
    return pd.DataFrame(CLINICS)


COST_SCALE = 0.65  # калибровка себестоимости под реалистичную маржу частной клиники


def build_services() -> pd.DataFrame:
    rows = []
    for i, (d, name, price, mat, ext, dur) in enumerate(SERVICES, start=1):
        mat, ext = round(mat * COST_SCALE), round(ext * COST_SCALE)
        rows.append(dict(
            service_id=f"S{i:03d}",
            direction_code=d,
            direction=DIRECTIONS[d],
            service_name=name,
            list_price=price,
            material_cost=mat,
            external_cost=ext,
            plan_cost=mat + ext,
            plan_margin_pct=round((price - mat - ext) / price, 4),
            duration_min=dur,
        ))
    return pd.DataFrame(rows)


def build_doctors(rng: random.Random) -> pd.DataFrame:
    rows, n = [], 0
    used_names = set()
    grades, probs = zip(*GRADE_PROBS)
    for direction, by_clinic in DOCTORS_PLAN.items():
        for clinic_id, count in by_clinic.items():
            for _ in range(count):
                n += 1
                while True:
                    name = f"{rng.choice(SURNAMES)} {rng.choice(INITIALS)}{rng.choice(INITIALS)}"
                    if name not in used_names:
                        used_names.add(name)
                        break
                grade = rng.choices(grades, probs)[0]
                pct, weight = GRADES[grade]
                rows.append(dict(
                    doctor_id=f"D{n:03d}",
                    doctor_name=name,
                    direction_code=direction,
                    direction=DIRECTIONS[direction],
                    specialty=rng.choice(SPECIALTY_NAMES[direction]),
                    clinic_id=clinic_id,
                    grade=grade,
                    pay_pct=pct,                      # доля врача от выручки
                    popularity_weight=weight,         # используется генератором визитов
                    hire_date=f"{rng.randint(2015, 2025)}-{rng.randint(1, 12):02d}-01",
                ))
    return pd.DataFrame(rows)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    rng = random.Random(SEED)

    clinics = build_clinics()
    services = build_services()
    doctors = build_doctors(rng)

    clinics.to_csv(OUT / "clinics.csv", index=False, encoding="utf-8-sig")
    services.to_csv(OUT / "services.csv", index=False, encoding="utf-8-sig")
    doctors.to_csv(OUT / "doctors.csv", index=False, encoding="utf-8-sig")

    print(f"clinics: {len(clinics)}  services: {len(services)}  doctors: {len(doctors)}")
    print("\nВрачей по направлениям:")
    print(doctors.groupby("direction").size().to_string())
    print("\nСредняя плановая маржа по направлениям:")
    print((services.groupby("direction")["plan_margin_pct"].mean() * 100).round(1).to_string())


if __name__ == "__main__":
    main()

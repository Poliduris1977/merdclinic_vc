"""
Полная пересборка проекта: данные -> сверка -> исправления -> отчёты -> дашборд -> тесты.
Запуск из корня проекта:  python run_all.py            (всё)
                          python run_all.py --from 11   (с шага 11: только отчёты и дашборд)
Каждый шаг пишет лог в logs/<номер>_<имя>.log. При ошибке цепочка останавливается.
"""
import argparse
import os
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PY = sys.executable
STEPS = [
    ("generate_reference", ["src/generate/generate_reference.py"]),
    ("generate_operations", ["src/generate/generate_operations.py"]),
    ("generate_1c", ["src/generate/generate_1c.py"]),
    ("generate_budget", ["src/generate/generate_budget.py"]),
    ("inject_errors", ["src/generate/inject_errors.py"]),
    ("reconcile_v1", ["src/reconcile/reconcile.py"]),
    ("apply_corrections", ["src/reconcile/apply_corrections.py"]),
    ("reconcile_corrected", ["src/reconcile/reconcile.py", "--input", "corrected", "--out", "output_corrected"]),
    ("inject_errors_v2", ["src/generate/inject_errors_v2.py"]),
    ("reconcile_v2", ["src/reconcile/reconcile.py", "--input", "dirty_v2", "--out", "output_v2", "--key", "answer_key_v2"]),
    ("build_pnl", ["src/reports/build_pnl.py"]),
    ("build_planfact", ["src/reports/build_planfact.py"]),
    ("build_factor", ["src/reports/build_factor.py"]),
    ("build_doctors", ["src/reports/build_doctors.py"]),
    ("build_motivation", ["src/reports/build_motivation.py"]),
    ("build_cashflow", ["src/reports/build_cashflow.py"]),
    ("build_dashboard", ["src/reports/build_dashboard.py"]),
    ("test_agent_offline", ["tests/test_agent_offline.py"]),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="start", type=int, default=1, help="номер шага, с которого начать (1-%d)" % len(STEPS))
    a = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    env = {**os.environ, "PYTHONUTF8": "1"}
    (ROOT / "logs").mkdir(exist_ok=True)
    t_all = time.time()
    for i, (name, cmd) in enumerate(STEPS, 1):
        if i < a.start:
            continue
        t0 = time.time()
        r = subprocess.run([PY, *cmd], cwd=ROOT, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace")
        (ROOT / "logs" / f"{i:02d}_{name}.log").write_text(r.stdout + r.stderr, encoding="utf-8")
        status = "OK  " if r.returncode == 0 else "FAIL"
        print(f"{i:2d}/{len(STEPS)} {status} {name} ({time.time() - t0:.0f} с)", flush=True)
        if r.returncode:
            print((r.stdout + r.stderr)[-1500:])
            sys.exit(f"Остановлено на шаге {i}: {name}. Полный лог: logs/{i:02d}_{name}.log")
    print(f"\nГотово за {time.time() - t_all:.0f} с. Отчёты: data/reports/, дашборд: data/reports/Dashboard_2025.html")


if __name__ == "__main__":
    main()

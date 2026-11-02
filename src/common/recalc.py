"""
Пересчёт xlsx в LibreOffice, чтобы в книге сохранились значения формул.

openpyxl записывает формулы без результатов: Excel посчитает их при открытии,
но предпросмотр (телефон, почта, браузер) покажет пустые ячейки. Здесь книга
прогоняется через LibreOffice headless и перезаписывается уже с кэшем значений.
Оформление (графики, фильтры, условное форматирование, шрифты) сохраняется.

Если LibreOffice не найден, файл остаётся как есть, а скрипт печатает предупреждение.
"""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path


def find_soffice():
    """soffice в PATH, а на Windows ещё в стандартных папках установки (там его в PATH обычно нет)."""
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
        if base and (Path(base) / "LibreOffice" / "program" / "soffice.exe").exists():
            return str(Path(base) / "LibreOffice" / "program" / "soffice.exe")
    return None


def recalculate(path, timeout=300):
    path = Path(path)
    soffice = find_soffice()
    if not soffice:
        print(f"Предупреждение: LibreOffice не найден, {path.name} сохранён без кэша значений формул.")
        return False
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out"
        cmd = [soffice, f"-env:UserInstallation={(Path(tmp) / 'profile').as_uri()}", "--headless", "--convert-to",
               "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out), str(path)]
        try:
            subprocess.run(cmd, check=True, capture_output=True, timeout=timeout)
        except (subprocess.SubprocessError, OSError) as e:
            print(f"Предупреждение: пересчёт {path.name} не удался ({e}); файл оставлен без кэша значений.")
            return False
        res = out / path.name
        if not res.exists():
            print(f"Предупреждение: LibreOffice не вернул файл для {path.name}.")
            return False
        shutil.move(str(res), str(path))
    print(f"Пересчитано: {path.name} (значения формул сохранены)")
    return True

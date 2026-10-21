# -*- coding: utf-8 -*-
"""Парсер дат старта когорт из присланного файла (Excel или скриншот).

Таблица «График запусков по бизнес-группам» очень приватна и никогда не
открывается по ссылке — поэтому ОМС присылает файл в чат, а не даёт ссылку.
Здесь читаем даты из присланного файла.

Excel: берём **только** листы Upskill и Reskill (в них живут cohort_name вида
`<слаг>_cohort_<номер>` и колонка «Дата старта когорты»). Остальные 38 листов
(архивы, кураторы, Анонс и т.п.) игнорируем.

Скриншот: распознаём текст через OCR (tesseract + pytesseract). Работает хуже
Excel, поэтому Excel — основной путь, скриншот — запасной.

Использование:
    python3 automation/date_parser.py --slug math-for-da-ds --file input/dates/График запусков.xlsx
    python3 automation/date_parser.py --slug math-for-da-ds --file input/dates/скрин.png

Пишет конфиг: config/dates.json  {номер когорты: "ГГГГ-ММ-ДД"}
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATE_FILE = ROOT / "config" / "dates.json"
INPUT_DIR = ROOT / "input" / "dates"

# Названия эмодзи-листов могут отличаться (новые эмодзи добавляются),
# поэтому матчим по подстроке, без учёта регистра и эмодзи.
UPSKILL_SHEETS = ("upskill",)
RESKILL_SHEETS = ("reskill",)

# «cohort_» или пусто (когда в ячейке только «data-analyst_201» — такое бывает)
COHORT_RE = re.compile(r"^([\w-]+?)(?:_cohort)?_(\d{2,3})$")
DATE_RE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
SLUG_RE = re.compile(r"[\w-]+")

# Допустимый диапазон годов. В колонке C всегда именно «Дата старта когорты»,
# и в таблице есть когорты ещё с 2023-го — не отсекаем их. Верхняя граница
# просто отбрасывает случайный мусор (невозможные даты в будущем).
_YEAR_MIN = 2020
_YEAR_MAX = 2035


def _norm_date(value) -> str | None:
    """Любое значение из ячейки -> 'YYYY-MM-DD' или None."""
    if value is None:
        return None
    if isinstance(value, str):
        # "2024-04-25" | "2024-04-25 00:00:00" | "25.04.2024" | "25/04/2024"
        value = value.strip()
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", value) or re.match(
            r"^(\d{1,2})[./](\d{1,2})[./](\d{4})$", value
        )
        if not m:
            return None
        if m.lastindex == 3:
            y, m_, d = m.group(1), m.group(2), m.group(3)
        else:
            d, m_, y = m.group(1), m.group(2), m.group(3)
        year, month, day = int(y), int(m_), int(d)
    elif hasattr(value, "year"):  # datetime / pandas.Timestamp
        year, month, day = value.year, value.month, value.day
    else:
        return None
    if not (_YEAR_MIN <= year <= _YEAR_MAX):
        return None
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return None
    return f"{year:04d}-{month:02d}-{day:02d}"


def _parse_excel(path: Path, slug: str) -> dict[int, str]:
    """Читаем листы Upskill и Reskill из .xlsx."""
    try:
        import pandas as pd
    except ImportError:
        print("openpyxl/pandas не найден. Поставь зависимости: pip install -r requirements.txt")
        sys.exit(1)

    found: dict[int, tuple[str, int]] = {}  # номер -> (дата, строка)

    # pd.read_excel с sheet_name=None читает ВСЕ 40 листов сразу — медленно
    # и лишние листы не нужны. Открываем вручную и читаем только нужные.
    try:
        import openpyxl
    except ImportError:
        print("openpyxl не найден. Поставь зависимости: pip install -r requirements.txt")
        sys.exit(1)

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    wanted = [name for name in wb.sheetnames
              if name.lower().strip() in UPSKILL_SHEETS + RESKILL_SHEETS]
    if not wanted:
        print(f"В файле нет листов Upskill/Reskill. Есть только: {wb.sheetnames}")
        sys.exit(1)
    print("Читаю листы:", ", ".join(wanted))

    for sheet_name in wanted:
        ws = wb[sheet_name]
        for rownum, row in enumerate(ws.iter_rows(min_row=3), start=3):
            col_name = row[1].value  # B — cohort_name
            col_date = row[2].value  # C — Дата старта когорты
            if not isinstance(col_name, str):
                continue
            name = col_name.strip()
            date = _norm_date(col_date)
            # Матчим только когорты указанного слага
            m = re.match(rf"^{re.escape(slug)}(?:_cohort)?_(\d+)$", name)
            if not m:
                continue
            num = int(m.group(1))
            if date is None:
                print(f"  ⚠ строка {rownum}: {name} — дату старта не распознал, пропускаю")
                continue
            # Если есть дубли (строки дублируются в разных листах) — берём первый
            # и помечаем повтор.
            if num in found:
                print(f"  ⚠ строка {rownum}: {name} уже была (дата {found[num][0]}), пропускаю дубль")
                continue
            found[num] = (date, rownum)

    wb.close()
    return {k: v[0] for k, v in found.items()}


# ---------------------------------------------------------------------------
# OCR-путь (скриншот таблицы) — запасной. Excel надёжнее, используй его.
# ---------------------------------------------------------------------------
def _parse_image(path: Path, slug: str) -> dict[int, str]:
    """Распознаём даты со скриншота таблицы через tesseract.

    Работает хуже Excel и зависит от качества скриншота. Excel всегда
    надёжнее — это основной путь, скриншот только запасной.
    """
    try:
        from PIL import Image
        import pytesseract
    except ImportError:
        print("Для скриншотов нужны PIL и pytesseract. Поставь: pip install -r requirements.txt")
        print("(и сам tesseract: на macOS — brew install tesseract, на Windows — https://digi.bib.uni-mannheim.de/tesseract/)")
        sys.exit(1)

    cmd = shutil.which("tesseract")
    if cmd:
        pytesseract.pytesseract.tesseract_cmd = cmd

    img = Image.open(path)
    ratio = 4000 / max(img.size)  # tesseract точнее на крупном изображении
    if ratio < 1:
        img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.LANCZOS)
    text = pytesseract.image_to_string(img, lang="rus+eng")

    # Имя когорты в скриншоте почти всегда искажается: дефисы/подчёркивания
    # теряются («math for da ds cohort 201»). Ловим по первым словам слага
    # и отдельно ищем число и даты в строке.
    words = [w for w in slug.split("-") if w]
    slug_prefix = r"[^0-9]".join(words)  # «math for da ds»

    found: dict[int, str] = {}
    for line in text.splitlines():
        # Дата ПЕРВОГО когорты - это «Дата старта». В строке с когортой дат
        # обычно несколько: старт, мягкий дедлайн, дедлайн добора.
        if not re.search(slug_prefix, line):
            continue
        m_num = re.search(r"(?:cohort|когорт)[_\s]?(\d{2,3})", line, re.IGNORECASE)
        if not m_num:
            continue
        # Все даты в строке; дата старта — последняя по порядку в строке
        # (мягкий дедлайн и дедлайн добора идут раньше старта).
        matches = DATE_RE.findall(line)
        if not matches:
            continue
        dates = sorted(f"{y}-{m}-{d}" for y, m, d in matches)
        num = int(m_num.group(1))
        if num == 0:
            continue
        found.setdefault(num, dates[-1])  # старт — самая поздняя дата в строке
    return found


def main() -> None:
    ap = argparse.ArgumentParser(description="Парсер дат старта когорт")
    ap.add_argument("--slug", required=True, help="Слаг профессии, например math-for-da-ds")
    ap.add_argument("--file", required=True, help="Путь к файлу: Excel или скриншот")
    args = ap.parse_args()

    path = Path(args.file)
    if not path.is_absolute() and not path.exists():
        # если пути нет — ищем в input/dates/ (туда ОМС кладёт файлы)
        candidate = INPUT_DIR / path
        if candidate.exists():
            path = candidate
    if not path.exists():
        print(f"Файл не найден: {path}")
        sys.exit(1)

    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        dates = _parse_excel(path, args.slug)
    elif ext in (".png", ".jpg", ".jpeg", ".bmp", ".webp"):
        dates = _parse_image(path, args.slug)
    else:
        print(f"Неподдерживаемый формат: {ext}. Пришли Excel (.xlsx) или скриншот (.png/.jpg).")
        sys.exit(1)

    if not dates:
        print(f"Не нашёл когорт слага «{args.slug}» в файле. "
              f"Проверь слаг (пример: math-for-da-ds, data-analyst).")
        sys.exit(1)

    # Сортируем по номеру, пишем в config/dates.json
    ordered = {str(k): v for k, v in sorted(dates.items())}
    DATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    DATE_FILE.write_text(json.dumps(ordered, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    nums = list(ordered.keys())
    print(f"\nНайдено когорт: {len(nums)} ({nums[0]}…{nums[-1]})")
    print(f"Записано в {DATE_FILE}")
    print("Первые 5:", json.dumps(dict(list(ordered.items())[:5]), ensure_ascii=False))
    print("Последние 5:", json.dumps(dict(list(ordered.items())[-5:]), ensure_ascii=False))


if __name__ == "__main__":
    main()
"""
Заведение когорт by OMS — парсинг таблицы «График запусков по бизнес группам».

Извлекает словарь «номер когорты → дата старта» для заданного слага из:
  - Excel (.xlsx/.xls) — лист с колонками вида <слаг>_cohort_<номер>;
  - скриншота (.png/.jpg) — колонки-строки с названием когорты и серией дат
    (по умолчанию используются tesseract — ocr).

Вход : input/dates/<любой файл>
Выход: config/dates.json  {"<номер>": "YYYY-MM-DD"}

Как ассистент это использует:
  1) кладёт файл от человека в input/dates/;
  2) запускает  python3 automation/date_parser.py --slug <слаг> --file <имя>;
  3) показывает человеку найденные даты и просит подтвердить/поправить.
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DATES_OUT = ROOT / "config" / "dates.json"

COHORT_RE = re.compile(r"(\d{3})")  # номер когорты = трёхзначное число в имени

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
SHEET_EXT = {".xlsx", ".xls", ".xlsm", ".csv"}


def _norm(value) -> str | None:
    """Валидная дата YYYY-MM-DD или None."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    # "2026-02-10", "10.02.2026", "10/02/2026"
    for fmt in (r"(\d{4})-(\d{2})-(\d{2})",
                r"(\d{1,2})[./](\d{1,2})[./](\d{4})",
                r"(\d{4})[./](\d{1,2})[./](\d{1,2})"):
        m = re.match(fmt, s)
        if m:
            a, b, c = m.groups()
            if len(a) == 4:
                return f"{a}-{int(b):02d}-{int(c):02d}"
            return f"{int(c):04d}-{int(b):02d}-{int(a):02d}"
    return None


def _parse_excel(path: Path, slug: str) -> dict[str, str]:
    """Ищем строки/колонки, где встречается '<slug>_cohort_<NNN>', и берём дату рядом."""
    df = pd.read_excel(path, sheet_name=None)
    out: dict[str, str] = {}
    for sheet_name, frame in df.items():
        arr = frame.astype(str).to_numpy()
        for row in arr:
            row_text = " | ".join(row)
            if slug + "_cohort_" not in row_text:
                continue
            match = COHORT_RE.search(row_text)
            if not match:
                continue
            num = match.group(1)
            # дата — первое значение строки, похожее на дату
            for cell in row:
                d = _norm(cell)
                if d:
                    out[num] = d
                    break
    return out


def _parse_image(path: Path, slug: str, tesseract_cmd: str | None) -> dict[str, str]:
    """OCR скриншота таблицы: ищем строки с именем когорты и датами."""
    try:
        import pytesseract
        from PIL import Image
    except ImportError as e:
        raise SystemExit(f"Нет библиотек для OCR (pillow, pytesseract): {e}")

    try:
        import PIL.Image  # noqa
        import pytesseract as _t  # noqa
    except ImportError:
        pass

    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd or pytesseract.pytesseract.tesseract_cmd
    raw = pytesseract.image_to_string(Image.open(path), lang="rus+eng")
    out: dict[str, str] = {}
    for line in raw.splitlines():
        if slug + "_cohort_" not in line:
            continue
        num_m = COHORT_RE.search(line)
        date_m = re.search(r"\d{1,2}[./]\d{1,2}[./]\d{2,4}|\d{4}-\d{2}-\d{2}", line)
        if num_m and date_m:
            out[num_m.group(1)] = _norm(date_m.group(0)) or ""
    return {k: v for k, v in out.items() if v}


def main() -> None:
    ap = argparse.ArgumentParser(description="Парсинг таблицы запусков когорт")
    ap.add_argument("--slug", required=True, help="слаг курса, например math-for-da-ds")
    ap.add_argument("--file", required=True, help="имя файла в input/dates/")
    ap.add_argument("--tesseract", default=None, help="путь к tesseract (если не в PATH)")
    args = ap.parse_args()

    src = ROOT / "input" / "dates" / args.file
    if not src.exists():
        raise SystemExit(f"Нет файла: {src}")

    if src.suffix.lower() in SHEET_EXT:
        dates = _parse_excel(src, args.slug)
    elif src.suffix.lower() in IMAGE_EXT:
        dates = _parse_image(src, args.slug, args.tesseract)
    else:
        raise SystemExit(f"Не знаю формат: {src.suffix} — поддерживаются {SHEET_EXT | IMAGE_EXT}")

    if not dates:
        raise SystemExit(f"Не нашёл дат для слага '{args.slug}' в {src.name}")

    DATES_OUT.write_text(
        json.dumps(dates, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(f"Готово: {len(dates)} дат → {DATES_OUT}")
    for num in sorted(dates):
        print(f"  {args.slug}_cohort_{num}  →  {dates[num]}")


if __name__ == "__main__":
    main()
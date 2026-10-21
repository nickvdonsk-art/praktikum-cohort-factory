"""
Заведение когорт by OMS — забор дат старта из Google Sheets.

Ссылка на таблицу «График запусков по бизнес группам» обычно имеет вид
  https://docs.google.com/spreadsheets/d/<ID>/edit#gid=0
Скачиваем лист как CSV (экспорт без авторизации) и вытаскиваем для слага
словарь «номер когорты → дата старта».

ВАЖНО: таблица должна быть открыта доступом «по ссылке» (Share → «Все, у кого
есть ссылка» → Просмотр). Если она приватная — export вернёт 401, и мы честно
скажем об этом человеку.
"""
from __future__ import annotations

import io
import re
import urllib.request
from pathlib import Path

import csv

ROOT = Path(__file__).resolve().parent.parent
DATES_OUT = ROOT / "config" / "dates.json"

SHEET_ID_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9_-]+)")
COHORT_RE = re.compile(r"(\d{3})")
DATE_RE = re.compile(
    r"(\d{4})-(\d{2})-(\d{2})"
    r"|(\d{1,2})[./](\d{1,2})[./](\d{4})"
    r"|(\d{1,2})[./](\d{1,2})[./](\d{2})"
)


class SheetAccessError(Exception):
    """Таблица не открыта по ссылке (401)."""


def _sheet_id(sheet_url: str) -> str:
    m = SHEET_ID_RE.search(sheet_url)
    if not m:
        raise ValueError(f"Не похоже на ссылку Google Sheets: {sheet_url}")
    return m.group(1)


def _fetch_csv(sheet_id: str) -> str:
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid=0"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read().decode("utf-8-sig", errors="replace")
    except urllib.error.HTTPError as e:
        if e.code == 401:
            raise SheetAccessError(
                "Таблица не открыта для доступа по ссылке. Открой её: "
                "Share → «Все, у кого есть ссылка» → Просмотр."
            ) from e
        raise


def _norm_date(value: str) -> str | None:
    s = value.strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:  # уже YYYY-MM-DD
        y, mo, d = m.groups()
        return f"{y}-{mo}-{d}"
    m = re.match(r"^(\d{1,2})[./](\d{1,2})[./](\d{2,4})$", s)
    if m:
        d, mo, y = m.groups()
        if len(y) == 2:
            y = "20" + y if int(y) <= 40 else "19" + y
        return f"{y}-{int(mo):02d}-{int(d):02d}"
    return None


def parse_dates_csv(text: str, slug: str) -> dict[str, str]:
    """Сканирует CSV-таблицу: имя когорты <slug>_cohort_<NNN> и дата рядом
    (в той же строке или в той же колонке). Возвращает {номер: YYYY-MM-DD}."""
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        return {}

    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)

    # 1. Индексы ячеек с именем когорты
    cells = []  # (r, c, номер)
    for r, row in enumerate(rows):
        for c, cell in enumerate(row):
            if slug + "_cohort_" in cell:
                num_m = COHORT_RE.search(cell)
                if num_m:
                    cells.append((r, c, num_m.group(1)))

    out: dict[str, str] = {}

    def _find_date_near(cell_r: int, cell_c: int) -> str | None:
        # по строке: влево и вправо
        row = rows[cell_r]
        for c in range(cell_c + 1, min(cell_c + 15, len(row))):
            d = _norm_date(row[c])
            if d:
                return d
        for c in range(max(0, cell_c - 15), cell_c):
            d = _norm_date(row[c])
            if d:
                return d
        # по колонке: вниз и вверх
        for r in range(cell_r + 1, min(cell_r + 15, n_rows)):
            if cell_c < len(rows[r]):
                d = _norm_date(rows[r][cell_c])
                if d:
                    return d
        for r in range(max(0, cell_r - 15), cell_r):
            if cell_c < len(rows[r]):
                d = _norm_date(rows[r][cell_c])
                if d:
                    return d
        return None

    for r, c, num in cells:
        d = _find_date_near(r, c)
        if d and num not in out:
            out[num] = d
    return out


def fetch_dates(sheet_url: str, slug: str) -> dict[str, str]:
    """Полный путь: ссылка → CSV → словарь «номер → дата». Пишет config/dates.json."""
    sid = _sheet_id(sheet_url)
    text = _fetch_csv(sid)
    dates = parse_dates_csv(text, slug)
    if not dates:
        raise ValueError(f"В таблице не нашлось дат для слага '{slug}' (ищи строки вида {slug}_cohort_<номер>).")
    DATES_OUT.write_text(json_dumps(dates), encoding="utf-8")
    return dates


def json_dumps(dates: dict[str, str]) -> str:
    import json

    return json.dumps(dates, ensure_ascii=False, indent=2) + "\n"


if __name__ == "__main__":
    import argparse
    import sys

    ap = argparse.ArgumentParser(description="Забор дат из Google Sheets")
    ap.add_argument("--url", required=True, help="ссылка на таблицу")
    ap.add_argument("--slug", required=True, help="слаг курса, например math-for-da-ds")
    args = ap.parse_args()
    try:
        got = fetch_dates(args.url, args.slug)
    except SheetAccessError as e:
        print(f"Ошибка доступа: {e}", file=sys.stderr)
        sys.exit(2)
    print(f"Готово: {len(got)} дат → {DATES_OUT}")
    for num in sorted(got):
        print(f"  {args.slug}_cohort_{num}  →  {got[num]}")
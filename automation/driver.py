"""
Заведение когорт by OMS — Playwright-движок (и selftest).

Основная работа идёт через automation/server.py (локальный сервис + инкогнито-окно).
Здесь — тонкий слой проверки окружения: есть ли Playwright, виден ли Chromium,
сходится ли config/steps.json с формами (только синтаксис — без входа в админку).

Запуск:  python3 automation/driver.py --selftest
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def selftest() -> int:
    problems: list[str] = []

    # 1. config/steps.json валиден и содержит поля из HTML модалки.
    steps = json.loads((ROOT / "config" / "steps.json").read_text(encoding="utf-8"))
    required_fields = {"name", "squad_number", "squad_type", "cohort", "track",
                       "curators_ids", "date_begin", "is_open",
                       "auto_compute_deadline_dates", "auto_compute_date_end",
                       "subscriptions_limit", "calendar_layer_id",
                       "is_new_documents_flow"}
    missing = required_fields - set(steps.get("fields", {}))
    if missing:
        problems.append(f"steps.json: нет полей {sorted(missing)}")
    for btn in ("add_cohort", "save", "cancel"):
        if btn not in steps.get("buttons", {}):
            problems.append(f"steps.json: нет кнопки '{btn}'")

    # 2. Playwright установлен и Chromium на месте.
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        problems.append(f"playwright не установлен: {e}")
        _print(problems)
        return 1

    with sync_playwright() as pw:
        try:
            pw.chromium.launch(headless=True)
        except Exception as e:  # pragma: no cover
            problems.append(f"Chromium не запускается: {e}\n  → выполни: playwright install chromium")
        else:
            print("OK: browser engine launch (headless)")

    _print(problems)
    return 1 if problems else 0


def _print(problems: list[str]) -> None:
    if problems:
        print("Проблемы:")
        for p in problems:
            print("  -", p)
    else:
        print("Все проверки пройдены.")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        sys.exit(selftest())
    print(__doc__)
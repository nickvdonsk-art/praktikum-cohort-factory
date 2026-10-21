#!/usr/bin/env bash
# Запуск заводилки: поднимает локальный сервис с инкогнито-окном админки.
# Используется ассистентом и человеком одинаково.
set -euo pipefail
cd "$(dirname "$0")/.."

if [ ! -d .venv ]; then
  echo "Нет .venv — сначала: python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt && playwright install chromium"
  exit 1
fi
source .venv/bin/activate

PYTHONPATH=. python3 -m automation.server "$@"
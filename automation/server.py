"""
Заведение когорт by OMS — локальный сервис автоматизации.

Держит браузерное инкогнито-окно админки admin.praktikum.yandex-team.ru
и отдаёт ассистенту (нейросети) пошаговые команды: snapshot DOM,
клик, ввод, выбор из дропдауна, заполнение формы по эталону.

Запуск:  bash automation/run.sh
API:     http://127.0.0.1:8742/docs  (Swagger) — только localhost.
"""
from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from playwright.sync_api import Page, TimeoutError as PWTimeout

ROOT = Path(__file__).resolve().parent.parent
CONFIG = ROOT / "config"
STEPS_JSON = CONFIG / "steps.json"
TEMPLATE_JSON = CONFIG / "template.json"
PORT = int(os.environ.get("OMSPORT", "8742"))
ADMIN_URL = "https://admin.praktikum.yandex-team.ru/squads-faculties/"

STEPS = json.loads(STEPS_JSON.read_text(encoding="utf-8"))

app = FastAPI(title="Заведение когорт by OMS", version="0.2.0")
app.add_middleware(
    CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"]
)

# --- состояние браузера -------------------------------------------------------
_browser = None
_context = None
_page: Page | None = None


# --- модели --------------------------------------------------------------------
class OpenRequest(BaseModel):
    url: str | None = None


class ClickRequest(BaseModel):
    selector: str = Field(description="CSS-селектор, например #name")
    description: str | None = None
    timeout_ms: int = 15000


class TypeRequest(BaseModel):
    selector: str
    value: str
    description: str | None = None


class SelectRequest(BaseModel):
    selector: str = Field(description="Триггер дропдауна, например #squad_type")
    option: str = Field(description="Текст пункта, например 'Стандартная'")
    description: str | None = None


class TickRequest(BaseModel):
    selector: str
    checked: bool | None = None  # None = переключить
    description: str | None = None


class CreateRequest(BaseModel):
    """Заполнить и сохранить одну когорту по эталону + дате старта."""
    cohort_number: int
    date_begin: str = Field(description="Дата старта в формате YYYY-MM-DD")
    template: dict = Field(default_factory=dict, description="Переопределения полей эталона")
    confirm_reference: bool = True


def _sel(name: str) -> str:
    """Селектор поля по имени из steps.json."""
    return STEPS["fields"][name]["selector"]


def _loc(page: Page, name_or_selector: str):
    """Принимает имя поля (из steps.json) или готовый CSS-селектор."""
    if name_or_selector in STEPS["fields"]:
        return page.locator(STEPS["fields"][name_or_selector]["selector"])
    return page.locator(name_or_selector)


# --- низкоуровневые помощники ---------------------------------------------------
def _active_page() -> Page:
    if _page is None:
        raise HTTPException(503, "Браузер не открыт. Вызови /api/open.")
    return _page


def _wait(page: Page, selector: str, timeout_ms: int) -> None:
    try:
        page.locator(selector).first.wait_for(state="visible", timeout=timeout_ms)
    except PWTimeout:
        raise HTTPException(404, f"Не найден элемент: {selector}")


def _visible_text(page: Page) -> str:
    """Текстовый 'снимок' страницы для навигации ассистента."""
    chunks: list[str] = []
    loc = page.locator("h1, h2, h3, .modal-title, .card-header, button, "
                       "label, input[type=text], input[type=number], "
                       "input[type=date], input[type=checkbox]")
    n = loc.count()
    for i in range(n):
        e = loc.nth(i)
        tag = e.evaluate("el => el.tagName.toLowerCase()")
        if tag == "input":
            v = e.input_value()
            checked = e.is_checked() if e.get_attribute("type") == "checkbox" else None
            tid = e.get_attribute("data-test-id") or ""
            out = f"input[{tid or e.get_attribute('id') or ''}]={v!r}"
            if checked is not None:
                out += f" checked={checked}"
            chunks.append(out)
        else:
            txt = e.inner_text().strip()
            if txt:
                chunks.append(f"<{tag}> {txt}")
    return "\n".join(chunks)


def _dropdown_select(page: Page, trigger: str, option: str, timeout_ms: int = 15000) -> None:
    """Открыть дропдаун по триггеру и выбрать пункт по тексту."""
    _wait(page, trigger, timeout_ms)
    page.locator(trigger).first.click()
    item = page.locator(
        f".dropdown__item:visible, [role=option]:visible, .prisma-combobox__option:visible"
    ).filter(has_text=option)
    try:
        item.first.click(timeout=timeout_ms)
    except PWTimeout:
        page.locator(f".dropdown__menu:visible >> text='{option}'").first.click(timeout=timeout_ms)


def _checkbox(page: Page, name_or_selector: str) -> bool:
    try:
        return _loc(page, name_or_selector).first.is_checked()
    except PWTimeout:
        return False


# --- логика одной когорты --------------------------------------------------------

def _open_cohort_form(page: Page, slug: str | None = None) -> None:
    """Открыть форму создания когорты (страница уже должна быть на /squads-faculties)."""
    if slug:
        _wait(page, STEPS["controls"]["profession_search"]["selector"], 15000)
        search = page.locator(STEPS["controls"]["profession_search"]["selector"]).first
        search.fill(slug)
        page.wait_for_timeout(700)
    _wait(page, STEPS["buttons"]["add_cohort"]["selector"], 15000)
    page.locator(STEPS["buttons"]["add_cohort"]["selector"]).first.click()
    _wait(page, "div[data-test-id='edit-card']", 15000)


def _read_current_form(page: Page) -> dict:
    """Прочитать значения полей открытой формы — используется как эталон на лету."""
    fields = STEPS["fields"]
    out: dict = {}

    def _v(key: str):
        try:
            return fields[key]["kind"]
        except KeyError:
            return "text"

    for key, spec in fields.items():
        if key in ("date_begin", "lite_payment_deadline", "payment_deadline", "date_end"):
            continue  # даты старта/дедлайнов в эталон не тащим
        kind = spec["kind"]
        sel = spec["selector"]
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=3000)
        except PWTimeout:
            continue
        if kind.startswith("checkbox"):
            out[key] = loc.is_checked()
        elif kind == "dropdown" or kind == "dropdown-multi":
            try:
                out[key] = page.locator(f"{sel} .dropdown__toggle, {sel} [role=combobox]").first.inner_text().strip()
            except Exception:
                out[key] = ""
        else:  # text, number, date…
            try:
                out[key] = loc.input_value()
            except Exception:
                out[key] = ""

    # отбрасываем пустые — их не ставим в эталон
    out = {k: v for k, v in out.items() if v not in ("", None)}
    # кураторы по требованию пользователя всегда пустые
    out.pop("curators_ids", None)
    return out


def _cohort_number(page: Page) -> int | None:
    try:
        v = page.locator("#squad_number").first.input_value()
        return int(v)
    except Exception:
        return None


# --- API ------------------------------------------------------------------------

@app.get("/api/status")
def status():
    global _page
    return {
        "alive": True,
        "browser_open": _page is not None,
        "url": _page.url if _page else None,
    }


@app.post("/api/open")
def open_browser(req: OpenRequest):
    global _browser, _context, _page
    from playwright.sync_api import sync_playwright

    if _browser is None:
        pw = sync_playwright().start()
        _browser = pw.chromium.launch(headless=False)
        _context = _browser.new_context()
    _page = _context.new_page()
    _page.goto(req.url or ADMIN_URL, wait_until="domcontentloaded")
    return {"ok": True, "url": _page.url}


@app.get("/api/snapshot")
def snapshot():
    page = _active_page()
    page.wait_for_timeout(600)
    return {"url": page.url, "text": _visible_text(page)}


@app.post("/api/click")
def click(req: ClickRequest):
    page = _active_page()
    _wait(page, req.selector, req.timeout_ms)
    page.locator(req.selector).first.click()
    return {"ok": True}


@app.post("/api/type")
def type_text(req: TypeRequest):
    page = _active_page()
    _wait(page, req.selector, 10000)
    loc = page.locator(req.selector).first
    loc.fill("")
    loc.fill(req.value)
    return {"ok": True}


@app.post("/api/select")
def select(req: SelectRequest):
    page = _active_page()
    _dropdown_select(page, req.selector, req.option)
    return {"ok": True}


@app.post("/api/tick")
def tick(req: TickRequest):
    page = _active_page()
    _wait(page, req.selector, 10000)
    loc = page.locator(req.selector).first
    if req.checked is None:
        loc.click()
    elif req.checked != loc.is_checked():
        loc.click()
    return {"ok": True, "checked": loc.is_checked()}


@app.post("/api/search_slug")
def search_slug(body: dict):
    """Ввести слаг в поиск факультетов и нажать «Добавить когорту»."""
    page = _active_page()
    slug = body.get("slug", "")
    if not slug:
        raise HTTPException(422, "Нужен slug")
    _open_cohort_form(page, slug)
    return {"ok": True, "cohort_prefill": _cohort_number(page)}


@app.post("/api/open_cohort_form")
def open_cohort_form(body: dict):
    page = _active_page()
    if body.get("slug"):
        _open_cohort_form(page, body["slug"])
    else:
        _open_cohort_form(page)
    return {"ok": True, "cohort_prefill": _cohort_number(page)}


@app.get("/api/current_form")
def current_form():
    """Прочитать заполненные поля открытой формы (после логина) — эталон на лету."""
    page = _active_page()
    return {"fields": _read_current_form(page)}


@app.post("/api/set_template_from_form")
def set_template_from_form(body: dict):
    """Взять значения открытой формы и записать в config/template.json.
    Кураторы всегда пустые (их заполняет человек при создании)."""
    page = _active_page()
    fields = _read_current_form(page)
    fields.pop("curators_ids", None)  # пользователь: всегда пусто
    merge = bool(body.get("merge", True))
    if merge and TEMPLATE_JSON.exists():
        try:
            cur = json.loads(TEMPLATE_JSON.read_text(encoding="utf-8"))
            cur.pop("curators_ids", None)
            cur.update(fields)
            fields = cur
        except Exception:
            pass
    fields["_comment"] = ("Эталон заполнения. Заполняется автоматически "
                          "чтением формы последней созданной когорты; кураторы всегда пустые.")
    TEMPLATE_JSON.write_text(json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"ok": True, "template": fields}


@app.post("/api/create_cohort")
def create_cohort(req: CreateRequest):
    """Полный сценарий на одну когорту: форма → заполнение → Сохранить."""
    page = _active_page()
    num = req.cohort_number
    date_begin = req.date_begin

    # форма уже открыта (ассистент сделал /api/open_cohort_form)
    _wait(page, "div[data-test-id='edit-card']", 15000)

    squad_type = req.template.get("squad_type", "Стандартная")
    _dropdown_select(page, _sel("squad_type"), squad_type)

    if req.template.get("track"):
        _dropdown_select(page, _sel("track"), req.template["track"])

    # кураторы НЕ заполняем — оставляем пустыми по требованию пользователя
    # (template.json не должен содержать curators_ids; если кто-то добавил — игнорируем)
    req.template.pop("curators_ids", None)

    for key in ("subscriptions_limit", "calendar_layer_id"):
        if req.template.get(key) is not None and key in STEPS["fields"]:
            _loc(page, key).first.fill(str(req.template[key]))

    loc = _loc(page, "date_begin").first
    loc.fill("")
    loc.fill(date_begin)

    for box in ("is_open", "is_new_documents_flow",
                "auto_compute_deadline_dates", "auto_compute_date_end"):
        want = req.template.get(box)
        if want is not None and box in STEPS["fields"]:
            loc = _loc(page, box).first
            try:
                if loc.is_checked() != want:
                    loc.click()
            except Exception:
                pass

    # Согласование имени/номера с диапазоном пользователя.
    actual_name = page.locator("#name").first.input_value()
    actual_num = _cohort_number(page)
    if actual_num and actual_num != num:
        # Форма подставила другой номер — честно сообщаем, не перезаписываем вслепую.
        page.screenshot(path=str(ROOT / "input" / "number_mismatch.png"), full_page=False)
        raise HTTPException(409,
                            f"Форма подставила номер {actual_num}, а ожидается {num}. "
                            f"Проверь: скриншот input/number_mismatch.png; название сейчас «{actual_name}»")

    # Сохранить.
    saver = page.locator(STEPS["buttons"]["save"]["selector"]).first
    try:
        saver.wait_for(state="visible", timeout=10000)
    except PWTimeout:
        raise HTTPException(500, "Кнопка «Сохранить» не найдена — возможно, форма не открылась.")
    if saver.is_disabled():
        page.screenshot(path=str(ROOT / "input" / "save_blocked.png"), full_page=False)
        raise HTTPException(422,
                            "Кнопка «Сохранить» заблокирована (валидация). "
                            "Скриншот: input/save_blocked.png")
    saver.click()
    page.wait_for_timeout(1200)

    return {"ok": True, "cohort_number": num, "date_begin": date_begin,
            "url": page.url,
            "note": "Проверь в админке, что когорта создана."}


@app.post("/api/screenshot")
def screenshot():
    page = _active_page()
    path = Path(tempfile.gettempdir()) / f"cohort_screen_{os.getpid()}.png"
    page.screenshot(path=str(path), full_page=False)
    return {"path": str(path)}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=PORT, log_level="info")
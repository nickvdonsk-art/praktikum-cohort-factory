"""
Заведение когорт by OMS — локальный сервис автоматизации.

Держит браузерное инкогнито-окно админки admin.praktikum.yandex-team.ru
и отдаёт ассистенту (нейросети) пошаговые команды: snapshot DOM,
клик, ввод, выбор из дропдауна, заполнение формы по эталону.

Запуск:  bash automation/run.sh
API:     http://127.0.0.1:8742/docs  (Swagger) — только localhost.
"""
from __future__ import annotations

import concurrent.futures
import functools
import json
import os
import queue
import tempfile
import threading
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

# Playwright sync нельзя звать из разных потоков: FastAPI исполняет эндпоинты
# в пуле потоков, и параллельный вызов из другого потока ломает greenlet
# («Cannot switch to a different thread»). Поэтому ВСЕ операции с браузером
# гоняем через один выделенный поток-исполнитель.
_pw_executor = concurrent.futures.ThreadPoolExecutor(max_workers=1,
                                                      thread_name_prefix="pw")


def _pw_serialized(fn):
    """Декоратор: выполняет тело эндпоинта в единственном pw-потоке."""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        fut = _pw_executor.submit(fn, *args, **kwargs)
        return fut.result()
    return wrapper


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
    """Заполнить и сохранить одну когорту по эталону + датам из Excel.

    date_end / lite_payment_deadline / payment_deadline — значения из Excel
    (колонки D / H / I). Используются как контроль: если форма автоподставляет
    другие значения (или не подставляет вовсе) — переопределяем их вручную.
    """
    cohort_number: int
    date_begin: str = Field(description="Дата старта (кол. C Excel) в формате YYYY-MM-DD")
    date_end: str | None = Field(default=None, description="Дата завершения обучения (кол. D Excel)")
    lite_payment_deadline: str | None = Field(default=None, description="Мягкий дедлайн (кол. H Excel)")
    payment_deadline: str | None = Field(default=None, description="Дедлайн добора (кол. I Excel)")
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
    """Открыть дропдаун по триггеру и выбрать пункт по тексту.

    Триггером м.б. контейнер поля (label[for=...]) — тогда сначала кликаем
    по его кнопке-тумблеру .dropdown__toggle.
    """
    _wait(page, trigger, timeout_ms)
    base = page.locator(trigger).first
    toggle = base.locator(":scope .dropdown__toggle")
    if toggle.count() > 0:
        toggle.first.click()
    else:
        base.click()
    page.wait_for_timeout(200)
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
    """Открыть форму создания когорты (страница уже должна быть на /squads-faculties).

    Поиск факультета — выпадающий список: слаг сначала вводится в поле поиска,
    затем нужный пункт (с текстом слага) выбирается из открывшегося списка.
    Без выбора пункта кнопка «Добавить когорту» недоступна.
    """
    if slug:
        _wait(page, STEPS["controls"]["profession_search"]["selector"], 15000)
        search = page.locator(STEPS["controls"]["profession_search"]["selector"]).first
        search.fill(slug)
        page.wait_for_timeout(400)
        # Список открывается не сразу после вставки — «дёргаем» поле:
        # стираем последний символ и печатаем его обратно (подсказка пользователя).
        search.press("Backspace")
        search.type(slug[-1])
        page.wait_for_timeout(800)  # дать списку открыться
        # выбрать пункт списка, содержащий слаг (по тексту, видимый)
        item = page.locator(
            ".prisma-combobox__option:visible, [role=option]:visible, "
            "li:visible, .dropdown__item:visible"
        ).filter(has_text=slug)
        try:
            item.first.click(timeout=8000)
        except PWTimeout:
            # fallback — точечно по тексту в любом элементе под поиском
            page.locator(f"text='{slug}'").first.click(timeout=5000)
        page.wait_for_timeout(1200)  # дождаться перехода на страницу факультета
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
@_pw_serialized
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
@_pw_serialized
def snapshot():
    page = _active_page()
    page.wait_for_timeout(600)
    return {"url": page.url, "text": _visible_text(page)}


@app.post("/api/click")
@_pw_serialized
def click(req: ClickRequest):
    page = _active_page()
    _wait(page, req.selector, req.timeout_ms)
    page.locator(req.selector).first.click()
    return {"ok": True}


@app.post("/api/type")
@_pw_serialized
def type_text(req: TypeRequest):
    page = _active_page()
    _wait(page, req.selector, 10000)
    loc = page.locator(req.selector).first
    loc.fill("")
    loc.fill(req.value)
    return {"ok": True}


@app.post("/api/select")
@_pw_serialized
def select(req: SelectRequest):
    page = _active_page()
    _dropdown_select(page, req.selector, req.option)
    return {"ok": True}


@app.post("/api/tick")
@_pw_serialized
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
@_pw_serialized
def search_slug(body: dict):
    """Ввести слаг в поиск факультетов и нажать «Добавить когорту»."""
    page = _active_page()
    slug = body.get("slug", "")
    if not slug:
        raise HTTPException(422, "Нужен slug")
    _open_cohort_form(page, slug)
    return {"ok": True, "cohort_prefill": _cohort_number(page)}


@app.post("/api/open_cohort_form")
@_pw_serialized
def open_cohort_form(body: dict):
    page = _active_page()
    if body.get("slug"):
        _open_cohort_form(page, body["slug"])
    else:
        _open_cohort_form(page)
    return {"ok": True, "cohort_prefill": _cohort_number(page)}


@app.get("/api/current_form")
@_pw_serialized
def current_form():
    """Прочитать заполненные поля открытой формы (после логина) — эталон на лету."""
    page = _active_page()
    return {"fields": _read_current_form(page)}


@app.post("/api/set_template_from_form")
@_pw_serialized
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
@_pw_serialized
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

    def _checkbox_value(name: str) -> bool | None:
        """Текущее состояние чекбокса (или None, если не нашли)."""
        if name not in STEPS["fields"]:
            return None
        try:
            return _loc(page, name).first.is_checked()
        except Exception:
            return None

    def _set_date_field(name: str, value: str | None) -> bool:
        """Вписать дату в поле (если оно разблокировано). Возвращает True, если поле заполнено."""
        if not value or name not in STEPS["fields"]:
            return False
        try:
            loc = _loc(page, name).first
            loc.fill("")
            loc.fill(value)
            return True
        except Exception:
            return False

    def _date_value(name: str) -> str:
        """Текущее значение поля-даты (что автоподставилось)."""
        try:
            return _loc(page, name).first.input_value().strip()
        except Exception:
            return ""

    # --- Проверка «Дата окончания» (date_end) ---
    # Если из Excel передана дата завершения обучения, а форма автоподставила
    # другую (или не подставила) — снимаем «вычислить автоматически» и вписываем вручную.
    # end_overridden запоминаем, чтобы в финальном блоке не включить автовычисление обратно.
    end_overridden = False
    need_end_manual = False
    if req.date_end:
        auto_end = _checkbox_value("auto_compute_date_end")
        cur_end = _date_value("date_end")
        if auto_end is not None and auto_end:
            if cur_end and cur_end != req.date_end:
                need_end_manual = True
            elif not cur_end:
                need_end_manual = True
        elif cur_end != req.date_end:
            need_end_manual = True
    if need_end_manual:
        box = "auto_compute_date_end"
        if box in STEPS["fields"]:
            loc = _loc(page, box).first
            try:
                if loc.is_checked():
                    loc.click()
            except Exception:
                pass
        _set_date_field("date_end", req.date_end)
        end_overridden = True

    # --- Проверка дедлайнов: мягкий + дедлайн добора ---
    # Если поля автозаполнились — ничего не делаем. Если пустые и переданы
    # значения из Excel — снимаем автовычисление и вписываем вручную.
    missing_deadlines = False
    for fname, xl_key in (("lite_payment_deadline", "lite_payment_deadline"),
                          ("payment_deadline", "payment_deadline")):
        if getattr(req, xl_key, None) and not _date_value(fname):
            missing_deadlines = True
    if missing_deadlines:
        box = "auto_compute_deadline_dates"
        if box in STEPS["fields"]:
            loc = _loc(page, box).first
            try:
                if loc.is_checked():
                    loc.click()
            except Exception:
                pass
        _set_date_field("lite_payment_deadline", req.lite_payment_deadline)
        _set_date_field("payment_deadline", req.payment_deadline)

    for box in ("is_open", "is_new_documents_flow",
                "auto_compute_deadline_dates", "auto_compute_date_end"):
        # Если дату окончания переопределили вручную — автовычисление не включаем.
        if box == "auto_compute_date_end" and end_overridden:
            continue
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

    return {"ok": True, "cohort_number": num,
            "date_begin": date_begin,
            "date_end": req.date_end,
            "lite_payment_deadline": req.lite_payment_deadline,
            "payment_deadline": req.payment_deadline,
            "date_end_overridden": end_overridden,
            "url": page.url,
            "note": "Проверь в админке, что когорта создана."}


@app.post("/api/dom")
@_pw_serialized
def dom_fragment(body: dict):
    """Вернуть html-фрагмент, содержащий искомую подстроку (например, номер когорты)."""
    page = _active_page()
    needle = (body.get("needle") or "").strip()
    if not needle:
        raise HTTPException(422, "Нужен needle")
    try:
        els = page.locator(f"text='{needle}'")
        n = els.count()
        rows = []
        for i in range(min(n, 6)):
            el = els.nth(i)
            html = el.evaluate("e => e.outerHTML")
            rows.append(html[:1500])
        return {"count": n, "htmls": rows}
    except Exception as e:
        raise HTTPException(500, str(e))


@app.post("/api/dom_modal")
@_pw_serialized
def dom_modal():
    """Вернуть HTML открытой модалки (для диагностики структуры формы)."""
    page = _active_page()
    sel = "div[data-test-id='edit-card']"
    try:
        loc = page.locator(sel).first
        loc.wait_for(state="visible", timeout=5000)
        html = loc.evaluate("e => e.outerHTML")
        return {"found": True, "selector": sel, "html": html[:6000]}
    except Exception:
        # fallback: собрать все edit-card в любом виде
        parts = []
        for sel2 in ("[class*='edit-card']", "[class*='editCard']", "[class*='modal']"):
            try:
                n = page.locator(sel2).count()
                if n:
                    first = page.locator(sel2).first
                    parts.append((sel2, n, first.evaluate("e => e.outerHTML")[:1500]))
            except Exception:
                pass
        return {"found": False, "candidates": parts}


@app.post("/api/hover")
@_pw_serialized
def hover(body: dict):
    """Навести мышь на элемент (по селектору или по тексту) — например, чтобы
    в строке когорты появилась кнопка «Редактировать»."""
    page = _active_page()
    selector = (body.get("selector") or "").strip()
    text = (body.get("text") or "").strip()
    if not selector and not text:
        raise HTTPException(422, "Нужен selector или text")
    try:
        if text:
            loc = page.get_by_text(text, exact=False).first
        else:
            loc = page.locator(selector).first
        loc.scroll_into_view_if_needed(timeout=8000)
        loc.hover(timeout=8000)
    except PWTimeout:
        raise HTTPException(404, f"Не найден элемент для hover: {selector or text}")
    page.wait_for_timeout(600)  # дать появиться кнопкам строки
    return {"ok": True, "target": text or selector}


@app.post("/api/screenshot")
@_pw_serialized
def screenshot():
    page = _active_page()
    path = Path(tempfile.gettempdir()) / f"cohort_screen_{os.getpid()}.png"
    page.screenshot(path=str(path), full_page=False)
    return {"path": str(path)}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=PORT, log_level="info")
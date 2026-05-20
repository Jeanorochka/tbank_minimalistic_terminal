# TradePanel_GUI.py
# GUI-скрипт для T-Bank Invest API:
# - выбор брокерского счета при запуске;
# - открытие LONG/SHORT лимиткой по текущей лучшей цене стакана;
# - SL% меняется прямо внутри окна;
# - TP% считается автоматически по RR 1:2;
# - по умолчанию SL = 1.3%, TP = 2.6% по RR 1:2;
# - ручное закрытие LONG/SHORT только рыночной заявкой;
# - автожурнал сделок в Excel: Дата / Капитал счёта / PnL / Тикер / Комментарии;
# - портфель внутри интерфейса;
# - SSL-проверка отключена жёстко для обхода CERTIFICATE_VERIFY_FAILED;
# - токен берётся из .env/.evn рядом со скриптом.
#
# Установка:
#   pip install requests openpyxl certifi
#
# Файл .env рядом со скриптом:
#   TINVEST_TOKEN=твой_токен
#
# Запуск:
#   python TradePanel_GUI.py

import os
import sys
import time
import uuid
import json
import threading
import warnings
import ctypes
from pathlib import Path
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

import requests
import certifi
from openpyxl import Workbook, load_workbook

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText

try:
    from urllib3.exceptions import InsecureRequestWarning
except Exception:
    InsecureRequestWarning = Warning


BASE_URL = "https://invest-public-api.tbank.ru/rest/tinkoff.public.invest.api.contract.v1"

APP_DIR = Path(__file__).resolve().parent
JOURNAL_PATH = APP_DIR / "trades_diary.xlsx"
STATE_PATH = APP_DIR / "active_trades.json"

ENV_FILE_USED: Path | None = None


def load_local_env() -> None:
    """Простая загрузка .env/.evn без внешних зависимостей.

    Файл должен лежать рядом со скриптом.
    Поддерживаются строки формата KEY=VALUE.
    """
    global ENV_FILE_USED

    # Основное имя — .env. На случай опечатки поддерживаем и .evn.
    candidates = [APP_DIR / ".env", APP_DIR / ".evn"]
    env_path = next((p for p in candidates if p.exists()), None)
    if not env_path:
        return

    ENV_FILE_USED = env_path

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()

        if not key:
            continue

        # Убираем кавычки, если пользователь написал TOKEN="..." или TOKEN='...'
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
            value = value[1:-1]

        # Значения из .env/.evn специально имеют приоритет над системными env,
        # чтобы не было путаницы со старыми токенами в PowerShell.
        os.environ[key] = value


load_local_env()

TOKEN = os.getenv("TINVEST_TOKEN") or os.getenv("INVEST_TOKEN")

# Жёсткий обход SSL-ошибки CERTIFICATE_VERIFY_FAILED.
# verify=False отключает проверку цепочки сертификатов в requests.
# Это менее безопасно, но решает проблему самоподписанного/подменённого сертификата.
SSL_VERIFY_MODE = "off_hardcoded"
SSL_VERIFY = False
warnings.simplefilter("ignore", InsecureRequestWarning)


def apply_theme(root: tk.Tk) -> None:
    root.configure(bg=BG)
    root.option_add("*Font", f"{FONT_FAMILY} 11")

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure(".", font=(FONT_FAMILY, 11), background=BG, foreground=FG)
    style.configure("TFrame", background=BG)
    style.configure("Panel.TFrame", background=PANEL_BG)

    style.configure("TLabel", background=BG, foreground=FG, font=(FONT_FAMILY, 11), anchor="center")
    style.configure("Title.TLabel", background=BG, foreground=FG, font=(FONT_FAMILY, 20, "bold"), anchor="center")
    style.configure("Muted.TLabel", background=BG, foreground=MUTED_FG, font=(FONT_FAMILY, 10), anchor="center")
    style.configure("Bold.TLabel", background=BG, foreground=FG, font=(FONT_FAMILY, 12, "bold"), anchor="center")

    style.configure(
        "TLabelframe",
        background=BG,
        foreground=FG,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        relief="solid",
    )
    style.configure("TLabelframe.Label", background=BG, foreground=FG, font=(FONT_FAMILY, 12, "bold"))

    style.configure(
        "TEntry",
        fieldbackground=INPUT_BG,
        foreground=INPUT_FG,
        insertcolor=INPUT_FG,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        padding=5,
    )
    style.configure(
        "TCombobox",
        fieldbackground=INPUT_BG,
        foreground=INPUT_FG,
        background=INPUT_BG,
        arrowcolor=INPUT_FG,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        padding=5,
    )
    style.map("TCombobox", fieldbackground=[("readonly", INPUT_BG)], foreground=[("readonly", INPUT_FG)])

    style.configure(
        "TButton",
        background=BUTTON_BG,
        foreground=FG,
        font=(FONT_FAMILY, 11, "bold"),
        borderwidth=1,
        relief="solid",
        focusthickness=1,
        focuscolor=FG,
        bordercolor=FG,
        lightcolor=FG,
        darkcolor=FG,
        padding=(12, 8),
    )
    style.map(
        "TButton",
        background=[("active", BUTTON_ACTIVE), ("pressed", BUTTON_ACTIVE)],
        foreground=[("active", FG), ("pressed", FG)],
    )

    style.configure("TSeparator", background=BORDER)


def style_textbox(widget: ScrolledText) -> None:
    widget.configure(
        bg=PANEL_BG_2,
        fg=FG,
        insertbackground=FG,
        selectbackground=BORDER,
        selectforeground=FG,
        font=(FONT_FAMILY, 10),
        relief="flat",
        borderwidth=0,
        padx=10,
        pady=8,
    )


# Визуальная тема. Цвет фона взят с присланной плашки: RGB(30, 44, 57).
BG = "#1E2C39"
PANEL_BG = "#263849"
PANEL_BG_2 = "#223342"
BUTTON_BG = "#111D29"      # очень тёмный, но не чисто чёрный
BUTTON_ACTIVE = "#182A3A"
FG = "#F2EDE3"   # костяной почти белый, мягче чистого #FFFFFF
MUTED_FG = "#D8D0C2"
INPUT_BG = "#F4F7FA"
INPUT_FG = "#101820"
BORDER = "#3B5064"
FONT_FAMILY = "Calibri"


def make_button(parent, text: str, command):
    """Кнопка на tk.Button: тёмная, костяной текст, тонкий ободок цвета текста."""
    return tk.Button(
        parent,
        text=text,
        command=command,
        bg=BUTTON_BG,
        fg=FG,
        activebackground=BUTTON_ACTIVE,
        activeforeground=FG,
        highlightbackground=FG,
        highlightcolor=FG,
        highlightthickness=1,
        bd=1,
        relief="solid",
        font=(FONT_FAMILY, 11, "bold"),
        padx=12,
        pady=7,
        cursor="hand2",
    )

# Риск-модель по умолчанию.
# Пользователь меняет только SL%. TP% автоматически = SL% * RR.
DEFAULT_SL_PERCENT = Decimal("0.013")    # 1.3% от цены входа
DEFAULT_RR = Decimal("2")                # RR 1:2 => TP = 2.6% при SL = 1.3%

ENTRY_WAIT_SECONDS = 12
OCO_CHECK_SECONDS = 3


def die(msg: str):
    raise RuntimeError(msg)


def headers():
    if not TOKEN:
        die("Не задан TINVEST_TOKEN. Создай файл .env рядом со скриптом и добавь строку: TINVEST_TOKEN=твой_токен")
    return {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json",
    }


def post(method: str, payload: dict) -> dict:
    url = f"{BASE_URL}.{method}"
    r = requests.post(
        url,
        headers=headers(),
        json=payload,
        timeout=20,
        verify=SSL_VERIFY,
    )
    if r.status_code >= 400:
        raise RuntimeError(f"API {method} вернул {r.status_code}: {r.text}")
    return r.json()


def q_to_decimal(q: dict | None) -> Decimal:
    if not q:
        return Decimal("0")
    units = Decimal(str(q.get("units", "0")))
    nano = Decimal(str(q.get("nano", 0))) / Decimal("1000000000")
    return units + nano


def decimal_to_q(value: Decimal) -> dict:
    value = Decimal(value)
    units = int(value)
    nano = int((value - Decimal(units)) * Decimal("1000000000"))
    return {"units": str(units), "nano": nano}


def money_to_decimal(m: dict | None) -> Decimal:
    return q_to_decimal(m)


def round_to_step(price: Decimal, step: Decimal) -> Decimal:
    if step <= 0:
        return price
    return (price / step).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * step


def fmt_dec(x: Decimal | str | int | float, places: int = 2) -> str:
    try:
        d = Decimal(str(x))
        return f"{d.quantize(Decimal('1.' + '0' * places))}"
    except Exception:
        return str(x)


def price_type_for(class_code: str) -> str:
    if str(class_code).upper() == "SPBFUT":
        return "PRICE_TYPE_POINT"
    return "PRICE_TYPE_CURRENCY"


# ---------- T-Invest API helpers ----------

def get_accounts() -> list[dict]:
    data = post("UsersService/GetAccounts", {})
    return data.get("accounts", [])


def get_portfolio(account_id: str) -> dict:
    return post("OperationsService/GetPortfolio", {"accountId": account_id})


def get_positions(account_id: str) -> dict:
    return post("OperationsService/GetPositions", {"accountId": account_id})


def find_instrument(ticker: str) -> dict:
    raw = ticker.strip().upper()
    preferred_class_code = None
    search_ticker = raw

    # Можно вводить: SBER, BRM6, BRM6_SPBFUT
    if "_" in raw:
        search_ticker, preferred_class_code = raw.split("_", 1)
        search_ticker = search_ticker.strip().upper()
        preferred_class_code = preferred_class_code.strip().upper()

    data = post("InstrumentsService/FindInstrument", {
        "query": search_ticker,
        "apiTradeAvailableFlag": True,
    })

    instruments = data.get("instruments", [])
    exact = []

    for inst in instruments:
        t = str(inst.get("ticker", "")).upper()
        cc = str(inst.get("classCode", "")).upper()

        if t != search_ticker:
            continue
        if preferred_class_code and cc != preferred_class_code:
            continue
        exact.append(inst)

    if not exact:
        die(
            f"Не нашёл торговый инструмент по тикеру {ticker}. "
            f"Для фьюча попробуй формат TICKER_CLASSCODE, например BRM6_SPBFUT."
        )

    exact.sort(key=lambda x: 0 if str(x.get("classCode", "")).upper() == "SPBFUT" else 1)
    return exact[0]


def get_instrument_id(inst: dict) -> str:
    instrument_id = inst.get("uid") or inst.get("instrumentUid") or inst.get("figi")
    if not instrument_id:
        die(f"У инструмента нет uid/figi. Ответ API: {inst}")
    return instrument_id


def get_instrument_full(instrument_id: str) -> dict:
    data = post("InstrumentsService/GetInstrumentBy", {
        "idType": "INSTRUMENT_ID_TYPE_UID",
        "id": instrument_id,
    })
    return data.get("instrument", {})


def get_min_step(instrument: dict, full: dict) -> Decimal:
    raw = full.get("minPriceIncrement") or instrument.get("minPriceIncrement")
    if raw:
        return q_to_decimal(raw)
    return Decimal("0.01")


def get_best_entry_price(instrument_id: str, side: str) -> Decimal:
    data = post("MarketDataService/GetOrderBook", {
        "instrumentId": instrument_id,
        "depth": 1,
    })
    bids = data.get("bids", [])
    asks = data.get("asks", [])

    if side == "BUY":
        if not asks:
            die("В стакане нет ask. Нельзя открыть LONG по текущей лимитной цене.")
        return q_to_decimal(asks[0]["price"])

    if side == "SELL":
        if not bids:
            die("В стакане нет bid. Нельзя открыть SHORT по текущей лимитной цене.")
        return q_to_decimal(bids[0]["price"])

    die("side должен быть BUY или SELL")


def calc_tp_sl(entry: Decimal, side: str, step: Decimal, sl_percent: Decimal, tp_percent: Decimal) -> tuple[Decimal, Decimal]:
    # sl_percent/tp_percent передаются в долях: 0.10 = 10%, 0.005 = 0.5%.
    # LONG: SL ниже, TP выше.
    # SHORT: SL выше, TP ниже.
    sl_distance = entry * sl_percent
    tp_distance = entry * tp_percent

    if side == "BUY":
        sl = entry - sl_distance
        tp = entry + tp_distance
    else:
        sl = entry + sl_distance
        tp = entry - tp_distance

    return round_to_step(tp, step), round_to_step(sl, step)


def post_limit_entry(account_id: str, instrument_id: str, qty: int, price: Decimal, side: str, class_code: str) -> str:
    direction = "ORDER_DIRECTION_BUY" if side == "BUY" else "ORDER_DIRECTION_SELL"
    payload = {
        "accountId": account_id,
        "instrumentId": instrument_id,
        "quantity": str(qty),
        "price": decimal_to_q(price),
        "direction": direction,
        "orderType": "ORDER_TYPE_LIMIT",
        "orderId": str(uuid.uuid4()),
        "timeInForce": "TIME_IN_FORCE_FILL_AND_KILL",
        "priceType": price_type_for(class_code),
        # Для открытия SHORT нужно согласие на сделку, которая может привести к непокрытой позиции.
        "confirmMarginTrade": True if side == "SELL" else False,
    }
    data = post("OrdersService/PostOrder", payload)
    order_id = data.get("orderId")
    if not order_id:
        die(f"Не получил orderId от API. Ответ: {data}")
    return order_id


def post_market_order(account_id: str, instrument_id: str, qty: int, direction: str, class_code: str, confirm_margin: bool = False) -> str:
    payload = {
        "accountId": account_id,
        "instrumentId": instrument_id,
        "quantity": str(qty),
        "price": decimal_to_q(Decimal("0")),
        "direction": direction,
        "orderType": "ORDER_TYPE_MARKET",
        "orderId": str(uuid.uuid4()),
        "timeInForce": "TIME_IN_FORCE_DAY",
        "priceType": price_type_for(class_code),
        "confirmMarginTrade": confirm_margin,
    }
    data = post("OrdersService/PostOrder", payload)
    order_id = data.get("orderId")
    if not order_id:
        die(f"Не получил orderId от API. Ответ: {data}")
    return order_id


def get_order_state(account_id: str, order_id: str) -> dict:
    return post("OrdersService/GetOrderState", {
        "accountId": account_id,
        "orderId": order_id,
    })


def wait_fill(account_id: str, order_id: str, timeout_sec: int = ENTRY_WAIT_SECONDS) -> int:
    last_state = None
    for _ in range(timeout_sec):
        state = get_order_state(account_id, order_id)
        last_state = state
        status = str(state.get("executionReportStatus", "")).upper()
        lots_executed = int(state.get("lotsExecuted", 0) or 0)

        if "FILL" in status and lots_executed > 0:
            return lots_executed
        if "REJECT" in status or "CANCEL" in status:
            return lots_executed
        time.sleep(1)

    return int((last_state or {}).get("lotsExecuted", 0) or 0)


def post_stop(account_id: str, instrument_id: str, qty: int, price: Decimal, original_side: str, stop_type: str, class_code: str) -> str:
    # Если открывали LONG, закрытие через SELL.
    # Если открывали SHORT, закрытие через BUY.
    exit_direction = "STOP_ORDER_DIRECTION_SELL" if original_side == "BUY" else "STOP_ORDER_DIRECTION_BUY"
    payload = {
        "accountId": account_id,
        "instrumentId": instrument_id,
        "quantity": str(qty),
        "price": decimal_to_q(price),
        "stopPrice": decimal_to_q(price),
        "direction": exit_direction,
        "expirationType": "STOP_ORDER_EXPIRATION_TYPE_GOOD_TILL_CANCEL",
        "stopOrderType": stop_type,
        "exchangeOrderType": "EXCHANGE_ORDER_TYPE_MARKET",
        "priceType": price_type_for(class_code),
        "orderId": str(uuid.uuid4()),
        "confirmMarginTrade": False,
    }
    data = post("StopOrdersService/PostStopOrder", payload)
    stop_id = data.get("stopOrderId")
    if not stop_id:
        die(f"Не получил stopOrderId. Ответ: {data}")
    return stop_id


def get_stop_orders(account_id: str) -> list[dict]:
    data = post("StopOrdersService/GetStopOrders", {"accountId": account_id})
    return data.get("stopOrders", [])


def cancel_stop_order(account_id: str, stop_order_id: str):
    if not stop_order_id:
        return
    post("StopOrdersService/CancelStopOrder", {
        "accountId": account_id,
        "stopOrderId": stop_order_id,
    })


# ---------- State / journal ----------

def load_state() -> dict:
    if not STATE_PATH.exists():
        return {"active_trades": []}
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {"active_trades": []}


def save_state(state: dict):
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def add_active_trade(trade: dict):
    state = load_state()
    trades = state.setdefault("active_trades", [])
    trades.append(trade)
    save_state(state)


def find_active_trade(account_id: str, ticker: str, side: str | None = None) -> dict | None:
    state = load_state()
    ticker = ticker.upper().strip()
    for trade in reversed(state.get("active_trades", [])):
        if trade.get("account_id") != account_id:
            continue
        if trade.get("ticker", "").upper() != ticker:
            continue
        if side and trade.get("side") != side:
            continue
        return trade
    return None


def remove_active_trade(trade_id: str):
    state = load_state()
    state["active_trades"] = [t for t in state.get("active_trades", []) if t.get("trade_id") != trade_id]
    save_state(state)


def ensure_journal():
    if JOURNAL_PATH.exists():
        wb = load_workbook(JOURNAL_PATH)
        ws = wb.active
        headers_row = [cell.value for cell in ws[1]]
        required = ["Дата", "Капитал счёта", "PnL", "Тикер", "Комментарии"]
        if headers_row[:5] != required:
            # Не ломаем существующий файл: просто создаем новый лист правильного формата.
            ws = wb.create_sheet("Trades")
            ws.append(required)
        wb.save(JOURNAL_PATH)
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Trades"
    ws.append(["Дата", "Капитал счёта", "PnL", "Тикер", "Комментарии"])
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 45
    wb.save(JOURNAL_PATH)


def append_journal_row(date_str: str, capital: Decimal | None, pnl: Decimal | None, ticker: str):
    ensure_journal()
    wb = load_workbook(JOURNAL_PATH)
    ws = wb.active
    if ws.max_row >= 1:
        first_row = [cell.value for cell in ws[1]]
        if first_row[:5] != ["Дата", "Капитал счёта", "PnL", "Тикер", "Комментарии"]:
            if "Trades" in wb.sheetnames:
                ws = wb["Trades"]
            else:
                ws = wb.create_sheet("Trades")
                ws.append(["Дата", "Капитал счёта", "PnL", "Тикер", "Комментарии"])

    ws.append([
        date_str,
        float(capital) if capital is not None else None,
        float(pnl) if pnl is not None else None,
        ticker,
        "",
    ])
    wb.save(JOURNAL_PATH)


def get_total_portfolio_value(account_id: str) -> Decimal:
    p = get_portfolio(account_id)
    return money_to_decimal(p.get("totalAmountPortfolio"))


# ---------- Window icon / taskbar ----------

def set_windows_app_user_model_id() -> None:
    """Заставляет Windows считать окно отдельным приложением, а не просто python.exe.

    Без этого иконка может поменяться в заголовке окна, но в панели задач
    останется стандартная иконка Python/Tkinter из-за группировки Windows.
    """
    try:
        if sys.platform.startswith("win"):
            app_id = "Kaiyah.TradePanel.1"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def find_app_icon() -> Path | None:
    candidates = [
        APP_DIR / "app_icon.ico",
        APP_DIR / "icon.ico",
        APP_DIR / "jeatrade.ico",
        APP_DIR / "app_icon.png",
        APP_DIR / "icon.png",
        APP_DIR / "jeatrade.png",
    ]
    return next((p for p in candidates if p.exists()), None)


def apply_window_icon(root: tk.Tk) -> None:
    icon_path = find_app_icon()
    if not icon_path:
        return

    try:
        if icon_path.suffix.lower() == ".ico":
            # iconbitmap сильнее влияет именно на Windows/taskbar.
            root.iconbitmap(default=str(icon_path))

        # iconphoto полезен для заголовка окна и fallback для png.
        if icon_path.suffix.lower() in {".png", ".gif"}:
            img = tk.PhotoImage(file=str(icon_path))
            root.iconphoto(True, img)
            root._app_icon_photo = img
        elif (APP_DIR / "app_icon.png").exists():
            img = tk.PhotoImage(file=str(APP_DIR / "app_icon.png"))
            root.iconphoto(True, img)
            root._app_icon_photo = img
    except Exception:
        pass



def remove_titlebar_text_and_icon(root: tk.Tk) -> None:
    """Оставляет обычное окно Windows, но убирает текст и маленькую иконку в верхней полосе.

    Важно: taskbar-иконка может оставаться app_icon.ico, а маленькая иконка
    в заголовке окна становится прозрачной. На не-Windows просто очищает title.
    """
    try:
        root.title("")
    except Exception:
        pass

    if sys.platform != "win32":
        return

    try:
        import ctypes

        root.update_idletasks()
        hwnd = int(root.winfo_id())

        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        WM_SETICON = 0x0080
        ICON_SMALL = 0
        ICON_BIG = 1
        ICON_SMALL2 = 2

        width = 16
        height = 16
        and_mask_size = (width * height) // 8
        xor_mask_size = width * height * 4

        and_mask = (ctypes.c_ubyte * and_mask_size)(*([0xFF] * and_mask_size))
        xor_mask = (ctypes.c_ubyte * xor_mask_size)(*([0x00] * xor_mask_size))

        hicon = user32.CreateIcon(
            kernel32.GetModuleHandleW(None),
            width,
            height,
            1,
            32,
            ctypes.byref(and_mask),
            ctypes.byref(xor_mask),
        )

        if hicon:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon)
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL2, hicon)
            # Не трогаем ICON_BIG, чтобы taskbar мог оставить нормальную app_icon.ico.
            root._blank_titlebar_hicon = hicon

        # Повторяем чуть позже: Windows/Tk иногда перерисовывает заголовок после маппинга окна.
        root.after(250, lambda: _remove_titlebar_icon_late(root))
    except Exception:
        # Fallback: хотя бы пустая PNG-иконка вместо пера Tkinter.
        try:
            blank = tk.PhotoImage(width=16, height=16)
            root.iconphoto(False, blank)
            root._blank_titlebar_photo = blank
        except Exception:
            pass


def _remove_titlebar_icon_late(root: tk.Tk) -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes
        hwnd = int(root.winfo_id())
        hicon = getattr(root, "_blank_titlebar_hicon", None)
        if hicon:
            WM_SETICON = 0x0080
            ICON_SMALL = 0
            ICON_SMALL2 = 2
            ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon)
            ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL2, hicon)
        root.title("")
    except Exception:
        pass

# ---------- GUI ----------

class JeatradeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        apply_theme(self.root)
        self.root.title("")
        remove_titlebar_text_and_icon(self.root)
        self.root.geometry("760x520")

        self.accounts: list[dict] = []
        self.account_id: str | None = None
        self.account_label: str = ""
        self.oco_started = False

        ensure_journal()
        self.build_account_window()
        self.load_accounts_async()

    def log(self, text: str):
        if hasattr(self, "log_box"):
            self.log_box.configure(state="normal")
            self.log_box.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        else:
            print(text)

    def run_async(self, target, on_done=None):
        def wrapper():
            try:
                result = target()
                if on_done:
                    self.root.after(0, lambda: on_done(result, None))
            except Exception as exc:
                if on_done:
                    self.root.after(0, lambda e=exc: on_done(None, e))
                else:
                    self.root.after(0, lambda e=exc: messagebox.showerror("Ошибка", str(e)))
        threading.Thread(target=wrapper, daemon=True).start()

    def build_account_window(self):
        for w in self.root.winfo_children():
            w.destroy()

        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Выбор брокерского счёта", style="Title.TLabel").pack(fill="x")
        ttk.Label(
            frame,
            text=f"SSL отключён. Env: {ENV_FILE_USED.name if ENV_FILE_USED else 'не найден'}. Журнал: {JOURNAL_PATH.name}",
            style="Muted.TLabel",
        ).pack(fill="x", pady=(6, 16))

        ttk.Label(frame, text="Брокерский счёт:", style="Bold.TLabel").pack(fill="x")
        self.account_var = tk.StringVar()
        self.account_combo = ttk.Combobox(frame, textvariable=self.account_var, state="readonly", width=95, justify="center")
        self.account_combo.pack(fill="x", pady=(6, 12))

        buttons = ttk.Frame(frame)
        buttons.pack(anchor="center")
        make_button(buttons, text="Обновить список счетов", command=self.load_accounts_async).pack(side="left", padx=6)
        make_button(buttons, text="Использовать выбранный счёт", command=self.select_account).pack(side="left", padx=6)

        self.status_label = ttk.Label(frame, text="Получаю список брокерских счетов...", style="Muted.TLabel")
        self.status_label.pack(fill="x", pady=(16, 6))

        self.log_box = ScrolledText(frame, height=16, state="disabled")
        style_textbox(self.log_box)
        self.log_box.pack(fill="both", expand=True, pady=(6, 0))

    def load_accounts_async(self):
        self.status_label.configure(text="Получаю список брокерских счетов...")
        self.log("Запрос UsersService/GetAccounts...")

        def task():
            return get_accounts()

        def done(result, error):
            if error:
                self.status_label.configure(text="Ошибка получения счетов")
                messagebox.showerror("Ошибка", str(error))
                self.log(str(error))
                return

            self.accounts = result or []
            if not self.accounts:
                self.status_label.configure(text="Счета не найдены")
                self.log("Счета не найдены.")
                return

            values = []
            for i, acc in enumerate(self.accounts, start=1):
                values.append(
                    f"{i}. {acc.get('name', '')} | id={acc.get('id')} | "
                    f"type={acc.get('type')} | status={acc.get('status')} | access={acc.get('accessLevel')}"
                )
            self.account_combo.configure(values=values)
            self.account_combo.current(0)
            self.status_label.configure(text=f"Найдено счетов: {len(values)}")
            self.log(f"Найдено счетов: {len(values)}")

        self.run_async(task, done)

    def select_account(self):
        idx = self.account_combo.current()
        if idx < 0 or idx >= len(self.accounts):
            messagebox.showwarning("Счёт", "Выбери счёт из списка.")
            return
        acc = self.accounts[idx]
        self.account_id = acc.get("id")
        self.account_label = f"{acc.get('name', '')} | {self.account_id}"
        if not self.account_id:
            messagebox.showerror("Ошибка", "У выбранного счёта нет id.")
            return
        self.build_trade_window()
        self.refresh_portfolio_async()
        self.start_oco_monitor()

    def build_trade_window(self):
        self.root.title("")
        remove_titlebar_text_and_icon(self.root)
        self.root.geometry("980x720")
        for w in self.root.winfo_children():
            w.destroy()

        main = ttk.Frame(self.root, padding=14)
        main.pack(fill="both", expand=True)

        top = ttk.Frame(main)
        top.pack(fill="x")
        ttk.Label(top, text="Торговля", style="Title.TLabel").pack(fill="x")
        ttk.Label(top, text=f"Счёт: {self.account_label}", style="Muted.TLabel").pack(fill="x", pady=(4, 8))
        make_button(top, text="Сменить счёт", command=self.build_account_window).pack(anchor="center")

        form = ttk.LabelFrame(main, text="Сделка", padding=14)
        form.pack(fill="x", pady=12)

        for col in range(8):
            form.grid_columnconfigure(col, weight=1)

        ttk.Label(form, text="Тикер:").grid(row=0, column=0, sticky="ew", padx=4)
        self.ticker_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.ticker_var, width=22, justify="center").grid(row=0, column=1, sticky="ew", padx=6)

        ttk.Label(form, text="Контракты/лоты:").grid(row=0, column=2, sticky="ew", padx=4)
        self.qty_var = tk.StringVar(value="1")
        ttk.Entry(form, textvariable=self.qty_var, width=12, justify="center").grid(row=0, column=3, sticky="ew", padx=6)

        ttk.Label(form, text="SL, %:").grid(row=0, column=4, sticky="ew", padx=4)
        self.sl_percent_var = tk.StringVar(value="1.3")
        self.sl_entry = ttk.Entry(form, textvariable=self.sl_percent_var, width=10, justify="center")
        self.sl_entry.grid(row=0, column=5, sticky="ew", padx=6)

        ttk.Label(form, text="TP авто:").grid(row=0, column=6, sticky="ew", padx=4)
        self.tp_preview_var = tk.StringVar(value="2.6%")
        ttk.Label(form, textvariable=self.tp_preview_var, style="Bold.TLabel").grid(row=0, column=7, sticky="ew", padx=6)

        self.rr_preview_var = tk.StringVar(value="SL 1.3% → TP 2.6%, RR 1:2")
        ttk.Label(
            form,
            textvariable=self.rr_preview_var,
            style="Muted.TLabel",
        ).grid(row=1, column=0, columnspan=8, sticky="ew", pady=(12, 0))

        self.position_status_var = tk.StringVar(value="Позиция по тикеру: —")
        ttk.Label(
            form,
            textvariable=self.position_status_var,
            style="Bold.TLabel",
        ).grid(row=2, column=0, columnspan=8, sticky="ew", pady=(10, 0))

        self.sl_percent_var.trace_add("write", lambda *_: self.update_tp_preview())
        self.update_tp_preview()

        buttons = ttk.Frame(form)
        buttons.grid(row=3, column=0, columnspan=8, pady=(14, 0))

        make_button(buttons, text="Открыть LONG", command=lambda: self.open_trade_async("BUY")).pack(side="left", padx=5)
        make_button(buttons, text="Открыть SHORT", command=lambda: self.open_trade_async("SELL")).pack(side="left", padx=5)
        ttk.Separator(buttons, orient="vertical").pack(side="left", fill="y", padx=12)
        make_button(buttons, text="Закрыть LONG по рынку", command=lambda: self.close_trade_async("BUY")).pack(side="left", padx=5)
        make_button(buttons, text="Закрыть SHORT по рынку", command=lambda: self.close_trade_async("SELL")).pack(side="left", padx=5)

        portfolio_frame = ttk.LabelFrame(main, text="Портфель", padding=12)
        portfolio_frame.pack(fill="both", expand=True, pady=(0, 12))

        portfolio_top = ttk.Frame(portfolio_frame)
        portfolio_top.pack(fill="x")
        self.portfolio_total_var = tk.StringVar(value="Капитал счёта: —")
        ttk.Label(portfolio_top, textvariable=self.portfolio_total_var, style="Bold.TLabel").pack(side="left")
        make_button(portfolio_top, text="Обновить портфель", command=self.refresh_portfolio_async).pack(side="right", padx=(6, 0))
        make_button(portfolio_top, text="Обновить позицию/PnL", command=self.refresh_current_position_async).pack(side="right", padx=(6, 0))

        self.portfolio_box = ScrolledText(portfolio_frame, height=14, state="disabled")
        style_textbox(self.portfolio_box)
        self.portfolio_box.pack(fill="both", expand=True, pady=(8, 0))

        log_frame = ttk.LabelFrame(main, text="Лог", padding=8)
        log_frame.pack(fill="both", expand=True)
        self.log_box = ScrolledText(log_frame, height=10, state="disabled")
        style_textbox(self.log_box)
        self.log_box.pack(fill="both", expand=True)

        self.log(f"Выбран счёт: {self.account_label}")
        self.log(f"Автожурнал: {JOURNAL_PATH}")

    def get_ticker_qty(self):
        ticker = self.ticker_var.get().strip().upper()
        if not ticker:
            die("Введи тикер.")
        try:
            qty = int(self.qty_var.get().strip())
        except ValueError:
            die("Количество должно быть целым числом.")
        if qty <= 0:
            die("Количество должно быть больше нуля.")
        return ticker, qty

    def parse_percent_from_window(self, raw_value: str, field_name: str) -> Decimal:
        raw = str(raw_value).strip().replace(",", ".").replace("%", "")
        if not raw:
            die(f"Поле {field_name} пустое.")
        try:
            value = Decimal(raw)
        except Exception:
            die(f"{field_name} должен быть числом. Пример: 10 или 0.5")
        if value <= 0:
            die(f"{field_name} должен быть больше нуля.")
        if value >= 100:
            die(f"{field_name} должен быть меньше 100%.")
        return value / Decimal("100")

    def update_tp_preview(self):
        try:
            sl_percent = self.parse_percent_from_window(self.sl_percent_var.get(), "SL, %")
            tp_percent = sl_percent * DEFAULT_RR
            self.tp_preview_var.set(self.format_percent(tp_percent))
            self.rr_preview_var.set(
                f"SL {self.format_percent(sl_percent)} → TP {self.format_percent(tp_percent)}, RR 1:{DEFAULT_RR.normalize()}"
            )
        except Exception:
            if hasattr(self, "tp_preview_var"):
                self.tp_preview_var.set("—")
            if hasattr(self, "rr_preview_var"):
                self.rr_preview_var.set("Введите SL, например 10 или 0.5")

    def get_risk_inputs(self) -> tuple[Decimal, Decimal]:
        sl_percent = self.parse_percent_from_window(self.sl_percent_var.get(), "SL, %")
        tp_percent = sl_percent * DEFAULT_RR
        return sl_percent, tp_percent

    def get_trade_inputs(self):
        ticker, qty = self.get_ticker_qty()
        sl_percent, tp_percent = self.get_risk_inputs()
        return ticker, qty, sl_percent, tp_percent

    def refresh_current_position_async(self):
        if not self.account_id:
            return

        ticker = self.ticker_var.get().strip().upper() if hasattr(self, "ticker_var") else ""
        if not ticker:
            if hasattr(self, "position_status_var"):
                self.position_status_var.set("Позиция по тикеру: введи тикер и нажми обновить")
            return

        def task():
            inst = find_instrument(ticker)
            instrument_id = get_instrument_id(inst)
            figi = inst.get("figi", "")
            real_ticker = inst.get("ticker", ticker)

            p = get_portfolio(self.account_id)
            for x in p.get("positions", []):
                pos_figi = x.get("figi", "")
                pos_uid = x.get("instrumentUid") or x.get("instrumentId") or x.get("uid")
                if pos_figi != figi and pos_uid != instrument_id:
                    continue

                qty = q_to_decimal(x.get("quantity"))
                avg = money_to_decimal(x.get("averagePositionPrice"))
                current = money_to_decimal(x.get("currentPrice"))
                pnl = money_to_decimal(x.get("expectedYield"))

                if qty > 0:
                    side_text = "LONG"
                elif qty < 0:
                    side_text = "SHORT"
                else:
                    side_text = "FLAT"

                return (
                    f"Позиция {real_ticker}: {side_text} | qty={qty} | "
                    f"avg={fmt_dec(avg)} | current={fmt_dec(current)} | PnL={fmt_dec(pnl)}"
                )

            return f"Позиция {real_ticker}: нет открытой позиции"

        def done(result, error):
            if error:
                self.position_status_var.set(f"Позиция/PnL: ошибка — {error}")
                self.log(f"Ошибка обновления позиции/PnL: {error}")
                return
            self.position_status_var.set(result)

        self.run_async(task, done)

    def format_percent(self, value: Decimal) -> str:
        percent = (value * Decimal("100")).normalize()
        return f"{percent}%"

    def open_trade_async(self, side: str):
        if not self.account_id:
            messagebox.showwarning("Счёт", "Сначала выбери счёт.")
            return
        try:
            ticker, qty, sl_percent, tp_percent = self.get_trade_inputs()
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))
            return

        side_name = "LONG" if side == "BUY" else "SHORT"
        rr_text = fmt_dec(tp_percent / sl_percent, 2)
        if not messagebox.askyesno(
            "Подтверждение",
            f"Открыть {side_name} {ticker}, количество {qty}?\n"
            f"SL: {self.format_percent(sl_percent)}; TP: {self.format_percent(tp_percent)}; RR: 1:{rr_text}"
        ):
            return

        self.log(
            f"Готовлю открытие {side_name}: {ticker}, qty={qty}, "
            f"SL={self.format_percent(sl_percent)}, TP={self.format_percent(tp_percent)}"
        )

        def task():
            before_portfolio = get_total_portfolio_value(self.account_id)
            inst = find_instrument(ticker)
            instrument_id = get_instrument_id(inst)
            class_code = inst.get("classCode", "")
            full = get_instrument_full(instrument_id)
            step = get_min_step(inst, full)
            entry = round_to_step(get_best_entry_price(instrument_id, side), step)
            tp, sl = calc_tp_sl(entry, side, step, sl_percent, tp_percent)

            order_id = post_limit_entry(self.account_id, instrument_id, qty, entry, side, class_code)
            filled_qty = wait_fill(self.account_id, order_id, ENTRY_WAIT_SECONDS)

            if filled_qty <= 0:
                return {
                    "status": "not_filled",
                    "ticker": ticker,
                    "entry": entry,
                    "order_id": order_id,
                }

            tp_id = post_stop(
                self.account_id,
                instrument_id,
                filled_qty,
                tp,
                side,
                "STOP_ORDER_TYPE_TAKE_PROFIT",
                class_code,
            )
            sl_id = post_stop(
                self.account_id,
                instrument_id,
                filled_qty,
                sl,
                side,
                "STOP_ORDER_TYPE_STOP_LOSS",
                class_code,
            )

            trade = {
                "trade_id": str(uuid.uuid4()),
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "account_id": self.account_id,
                "ticker": inst.get("ticker", ticker),
                "class_code": class_code,
                "instrument_id": instrument_id,
                "side": side,
                "qty": filled_qty,
                "entry_price": str(entry),
                "tp_price": str(tp),
                "sl_price": str(sl),
                "sl_percent": str(sl_percent),
                "tp_percent": str(tp_percent),
                "tp_stop_id": tp_id,
                "sl_stop_id": sl_id,
                "entry_portfolio": str(before_portfolio),
            }
            add_active_trade(trade)

            return {
                "status": "ok",
                "side": side,
                "side_name": side_name,
                "ticker": inst.get("ticker", ticker),
                "qty": filled_qty,
                "requested_qty": qty,
                "entry": entry,
                "tp": tp,
                "sl": sl,
                "tp_id": tp_id,
                "sl_id": sl_id,
                "trade": trade,
            }

        def done(result, error):
            if error:
                messagebox.showerror("Ошибка открытия", str(error))
                self.log(f"Ошибка открытия: {error}")
                return
            if result["status"] == "not_filled":
                self.log(f"Вход не исполнился: {result['ticker']} по {result['entry']}. TP/SL не ставились.")
                messagebox.showinfo("Не исполнено", "Входная лимитка не исполнилась. TP/SL не ставились.")
                self.refresh_portfolio_async()
                self.refresh_current_position_async()
                return
            self.log(
                f"Открыт {result['side_name']} {result['ticker']} qty={result['qty']} "
                f"entry={result['entry']} TP={result['tp']} SL={result['sl']}"
            )
            if result["qty"] < result["requested_qty"]:
                self.log(f"Частичное исполнение: {result['qty']} из {result['requested_qty']}.")
            messagebox.showinfo("Готово", f"Открыто: {result['side_name']} {result['ticker']}\nTP: {result['tp']}\nSL: {result['sl']}")
            self.refresh_portfolio_async()
            self.refresh_current_position_async()

        self.run_async(task, done)

    def close_trade_async(self, opened_side: str):
        # opened_side BUY = закрываем LONG продажей.
        # opened_side SELL = закрываем SHORT покупкой.
        if not self.account_id:
            messagebox.showwarning("Счёт", "Сначала выбери счёт.")
            return
        try:
            ticker, qty = self.get_ticker_qty()
        except Exception as exc:
            messagebox.showerror("Ошибка", str(exc))
            return

        position_name = "LONG" if opened_side == "BUY" else "SHORT"
        close_direction = "ORDER_DIRECTION_SELL" if opened_side == "BUY" else "ORDER_DIRECTION_BUY"

        if not messagebox.askyesno("Подтверждение", f"Закрыть {position_name} {ticker}, количество {qty}, по рынку?"):
            return

        self.log(f"Ручное закрытие {position_name}: {ticker}, qty={qty}, MARKET")

        def task():
            trade = find_active_trade(self.account_id, ticker, opened_side)

            inst = find_instrument(ticker)
            instrument_id = get_instrument_id(inst)
            class_code = inst.get("classCode", "")

            # Сначала снимаем защитные стопы, если они были записаны в state.
            if trade:
                for stop_id in (trade.get("tp_stop_id"), trade.get("sl_stop_id")):
                    try:
                        cancel_stop_order(self.account_id, stop_id)
                    except Exception:
                        # Если стоп уже исполнился/исчез, не валим весь ручной выход.
                        pass

            order_id = post_market_order(
                self.account_id,
                instrument_id,
                qty,
                close_direction,
                class_code,
                confirm_margin=False,
            )
            filled_qty = wait_fill(self.account_id, order_id, timeout_sec=20)

            after_portfolio = get_total_portfolio_value(self.account_id)
            pnl = None
            if trade and trade.get("entry_portfolio"):
                pnl = after_portfolio - Decimal(str(trade["entry_portfolio"]))
                remove_active_trade(trade["trade_id"])

            append_journal_row(
                datetime.now().strftime("%d.%m.%Y"),
                after_portfolio,
                pnl,
                inst.get("ticker", ticker),
            )

            return {
                "ticker": inst.get("ticker", ticker),
                "position_name": position_name,
                "filled_qty": filled_qty,
                "order_id": order_id,
                "capital": after_portfolio,
                "pnl": pnl,
                "journal": str(JOURNAL_PATH),
                "had_state": bool(trade),
            }

        def done(result, error):
            if error:
                messagebox.showerror("Ошибка закрытия", str(error))
                self.log(f"Ошибка закрытия: {error}")
                return
            self.log(
                f"Закрытие {result['position_name']} {result['ticker']}: исполнено={result['filled_qty']}, "
                f"капитал={fmt_dec(result['capital'])}, PnL={fmt_dec(result['pnl']) if result['pnl'] is not None else '—'}"
            )
            if not result["had_state"]:
                self.log("Внимание: активная сделка не найдена в active_trades.json, PnL в журнале может быть пустым/неточным.")
            messagebox.showinfo("Закрыто", f"Закрыто по рынку.\nЖурнал: {result['journal']}")
            self.refresh_portfolio_async()
            self.refresh_current_position_async()

        self.run_async(task, done)

    def refresh_portfolio_async(self):
        if not self.account_id:
            return

        def task():
            p = get_portfolio(self.account_id)
            pos = p.get("positions", [])
            total = money_to_decimal(p.get("totalAmountPortfolio"))
            expected_yield = money_to_decimal(p.get("expectedYield"))
            lines = []
            lines.append(f"Капитал счёта: {fmt_dec(total)}")
            lines.append(f"Ожидаемая доходность портфеля: {fmt_dec(expected_yield)}")
            lines.append("")
            lines.append("Позиции:")

            if not pos:
                lines.append("— позиций нет или API не вернул позиции")
            else:
                for x in pos:
                    figi = x.get("figi", "")
                    instrument_type = x.get("instrumentType", "")
                    qty = q_to_decimal(x.get("quantity"))
                    current_price = money_to_decimal(x.get("currentPrice"))
                    average_price = money_to_decimal(x.get("averagePositionPrice"))
                    expected = money_to_decimal(x.get("expectedYield"))
                    lines.append(
                        f"{instrument_type:10} {figi:14} qty={qty} avg={average_price} current={current_price} pnl={expected}"
                    )
            return total, "\n".join(lines)

        def done(result, error):
            if error:
                self.log(f"Ошибка обновления портфеля: {error}")
                return
            total, text = result
            self.portfolio_total_var.set(f"Капитал счёта: {fmt_dec(total)}")
            self.portfolio_box.configure(state="normal")
            self.portfolio_box.delete("1.0", "end")
            self.portfolio_box.insert("1.0", text)
            self.portfolio_box.configure(state="disabled")
            self.log("Портфель обновлён.")

        self.run_async(task, done)

    def start_oco_monitor(self):
        if self.oco_started:
            return
        self.oco_started = True

        def loop():
            while True:
                try:
                    if self.account_id:
                        self.check_oco_once()
                except Exception as exc:
                    self.root.after(0, lambda e=exc: self.log(f"OCO-monitor: {e}"))
                time.sleep(OCO_CHECK_SECONDS)

        threading.Thread(target=loop, daemon=True).start()

    def check_oco_once(self):
        state = load_state()
        trades = state.get("active_trades", [])
        if not trades or not self.account_id:
            return

        active_stops = get_stop_orders(self.account_id)
        active_ids = {x.get("stopOrderId") for x in active_stops}

        changed = False
        for trade in list(trades):
            if trade.get("account_id") != self.account_id:
                continue
            tp_id = trade.get("tp_stop_id")
            sl_id = trade.get("sl_stop_id")
            tp_active = tp_id in active_ids
            sl_active = sl_id in active_ids

            if tp_active and sl_active:
                continue

            # Один из защитных стопов исчез: считаем, что он исполнился/отменился, второй снимаем.
            if not tp_active and sl_active:
                try:
                    cancel_stop_order(self.account_id, sl_id)
                except Exception:
                    pass
                self.log(f"OCO: TP по {trade.get('ticker')} исчез/сработал, SL снят.")
            elif not sl_active and tp_active:
                try:
                    cancel_stop_order(self.account_id, tp_id)
                except Exception:
                    pass
                self.log(f"OCO: SL по {trade.get('ticker')} исчез/сработал, TP снят.")
            else:
                self.log(f"OCO: обе защитные заявки по {trade.get('ticker')} отсутствуют.")

            try:
                after_portfolio = get_total_portfolio_value(self.account_id)
                pnl = None
                if trade.get("entry_portfolio"):
                    pnl = after_portfolio - Decimal(str(trade["entry_portfolio"]))
                append_journal_row(
                    datetime.now().strftime("%d.%m.%Y"),
                    after_portfolio,
                    pnl,
                    trade.get("ticker", ""),
                )
            except Exception as exc:
                self.log(f"OCO: не смог записать журнал: {exc}")

            trades.remove(trade)
            changed = True

        if changed:
            state["active_trades"] = trades
            save_state(state)
            self.root.after(0, self.refresh_portfolio_async)


def main():
    # Важно: AppUserModelID надо задать ДО создания Tk-окна,
    # иначе панель задач Windows может оставить иконку python.exe.
    set_windows_app_user_model_id()

    root = tk.Tk()
    apply_window_icon(root)
    remove_titlebar_text_and_icon(root)

    app = JeatradeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

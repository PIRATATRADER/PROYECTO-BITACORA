"""
DAS Trader Montage Monitor + Ask Edgar Dilution Overlay
-------------------------------------------------------
Monitors the active DAS Trader montage window for ticker changes,
fetches dilution risk data from the Ask Edgar API, and displays
results in an always-on-top overlay panel.

Includes Top Gainers panel with real-time market data.
"""

import os
import threading
import time
import webbrowser
import requests
import tkinter as tk
import win32gui
import re
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

# Load .env file if python-dotenv is installed
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ── Config ──────────────────────────────────────────────────────────────────
# API keys – set these as environment variables or in a .env file
# See .env.example for details
ASKEDGAR_API_KEY = os.environ.get("ASKEDGAR_API_KEY", "")
TRANSLATE_ENABLED = os.environ.get("TRANSLATE", "FALSE").upper() == "TRUE"
SHOW_COMPANY_NAME = os.environ.get("SHOW_COMPANY_NAME", "FALSE").upper() == "TRUE"
CLAUDE_API_KEY = os.environ.get("CLAUDE_API_KEY", "")
API_LOG_ENABLED = os.environ.get("API_LOG", "FALSE").upper() == "TRUE"

# ── Translation helper ────────────────────────────────────────────────────
_translator = None
_translate_cache = {}  # text -> translated text
def translate_text(text: str) -> str:
    """Translate English text to Spanish. Cached in memory. Returns original on failure."""
    if not TRANSLATE_ENABLED or not text:
        return text
    if text in _translate_cache:
        return _translate_cache[text]
    try:
        global _translator
        if _translator is None:
            from deep_translator import GoogleTranslator
            _translator = GoogleTranslator(source='en', target='es')
        result = _translator.translate(text)
        _translate_cache[text] = result
        return result
    except Exception:
        return text

# ── Country flag helper ────────────────────────────────────────────────────
_flag_cache = {}  # country_code -> bytes or None (in-memory cache for this session)
_FLAGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "flags")

def fetch_flag_image(country_code: str) -> bytes | None:
    """Return flag PNG bytes. Reads from assets/flags/{code}.png if present,
    otherwise downloads from flagcdn.com and saves locally for future use."""
    if not country_code or len(country_code) != 2:
        return None
    code = country_code.lower()

    # 1. In-memory cache (fastest — already fetched this session)
    if code in _flag_cache:
        return _flag_cache[code]

    # 2. Local disk cache (fast — no network needed)
    local_path = os.path.join(_FLAGS_DIR, f"{code}.png")
    if os.path.exists(local_path):
        try:
            with open(local_path, "rb") as f:
                data = f.read()
            _flag_cache[code] = data
            return data
        except Exception:
            pass

    # 3. Download from CDN and save to disk for next time
    try:
        resp = requests.get(f"https://flagcdn.com/16x12/{code}.png", timeout=3)
        if resp.status_code == 200:
            data = resp.content
            _flag_cache[code] = data
            try:
                os.makedirs(_FLAGS_DIR, exist_ok=True)
                with open(local_path, "wb") as f:
                    f.write(data)
            except Exception:
                pass  # disk save failed, still return the bytes
            return data
    except Exception:
        pass

    _flag_cache[code] = None
    return None

# PhotoImage cache: one PhotoImage per country, reused across all gainer rows
# and the header flag. Avoids recreating PhotoImage textures on every rebuild.
_flag_photo_cache = {}  # country_code -> ImageTk.PhotoImage or None
def get_flag_photo(country_code: str):
    """Return a cached PhotoImage for a country flag. Returns None if unavailable."""
    if not country_code or len(country_code) != 2:
        return None
    code = country_code.lower()
    if code in _flag_photo_cache:
        return _flag_photo_cache[code]
    data = fetch_flag_image(code)
    if not data:
        _flag_photo_cache[code] = None
        return None
    try:
        from PIL import Image, ImageTk
        import io
        img = Image.open(io.BytesIO(data))
        photo = ImageTk.PhotoImage(img)
        _flag_photo_cache[code] = photo
        return photo
    except Exception:
        _flag_photo_cache[code] = None
        return None

# ── ET to Spain time conversion ───────────────────────────────────────────
def convert_api_date(raw_date: str) -> tuple[str, str]:
    """Convert raw API date (always EST/UTC-5) to Spain time.
    Returns (date_str, time_str) e.g. ('10ABR26', '13:00')"""
    if not raw_date or len(raw_date) < 16:
        return (raw_date, "")
    try:
        from datetime import datetime, timezone, timedelta
        from zoneinfo import ZoneInfo
        MONTHS_ES = {
            1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
            7: "JUL", 8: "AGO", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC"
        }
        date_str = raw_date[:16].replace("T", " ")
        naive = datetime.strptime(date_str, "%Y-%m-%d %H:%M")
        est_fixed = timezone(timedelta(hours=-5))
        utc_dt = naive.replace(tzinfo=est_fixed)
        spain_dt = utc_dt.astimezone(ZoneInfo("Europe/Madrid"))
        day = spain_dt.strftime("%d")
        month = MONTHS_ES[spain_dt.month]
        year = spain_dt.strftime("%y")
        date_display = f"{day}{month}{year}"
        time_display = spain_dt.strftime("%H:%M")
        return (date_display, time_display)
    except Exception:
        return (raw_date[:16].replace("T", " "), "")

# Spain flag image cached for reuse in feed dates
_spain_flag_photo = None
def get_spain_flag_photo():
    """Get Spain flag as PhotoImage, cached. Returns None if unavailable."""
    global _spain_flag_photo
    if _spain_flag_photo is not None:
        return _spain_flag_photo
    flag_bytes = fetch_flag_image("ES")
    if flag_bytes:
        try:
            from PIL import Image, ImageTk
            import io
            img = Image.open(io.BytesIO(flag_bytes))
            _spain_flag_photo = ImageTk.PhotoImage(img)
            return _spain_flag_photo
        except Exception:
            pass
    return None


# Notes icon for the header button (loaded once from assets/notes-icon.png)
_notes_icon_photo = None
_notes_icon_loaded = False
def get_notes_icon_photo():
    """Get notes icon as PhotoImage resized to 18x18, cached. Returns None if file missing."""
    global _notes_icon_photo, _notes_icon_loaded
    if _notes_icon_loaded:
        return _notes_icon_photo
    _notes_icon_loaded = True
    png_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "notes-icon.png")
    if not os.path.exists(png_path):
        return None
    try:
        from PIL import Image, ImageTk
        img = Image.open(png_path).convert("RGBA")
        img = img.resize((18, 18), Image.LANCZOS)
        _notes_icon_photo = ImageTk.PhotoImage(img)
        return _notes_icon_photo
    except Exception:
        return None

# "New" icon for highlighted news (loaded from assets/new-icon.png if present)
_new_icon_photo = None
_new_icon_loaded = False
def get_new_icon_photo():
    """Get 'new' icon as PhotoImage resized to 16x16, cached. Returns None if file missing."""
    global _new_icon_photo, _new_icon_loaded
    if _new_icon_loaded:
        return _new_icon_photo
    _new_icon_loaded = True
    png_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "new-icon.png")
    if not os.path.exists(png_path):
        return None
    try:
        from PIL import Image, ImageTk
        img = Image.open(png_path).convert("RGBA")
        img = img.resize((16, 16), Image.LANCZOS)
        _new_icon_photo = ImageTk.PhotoImage(img)
        return _new_icon_photo
    except Exception:
        return None

# ── Tooltip helper (for hover text on widgets like flag images) ───────────
def _attach_tooltip(widget, text: str):
    """Attach a simple hover tooltip to a widget. Replaces any existing tooltip."""
    if not text:
        return
    import tkinter as _tk
    tip = {"win": None}
    def _show(event):
        if tip["win"] is not None:
            return
        t = _tk.Toplevel(widget)
        t.overrideredirect(True)
        t.attributes("-topmost", True)
        _tk.Label(t, text=text, fg="white", bg="#333333",
                  font=("Segoe UI", 9), padx=6, pady=3).pack()
        t.geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
        tip["win"] = t
    def _hide(event):
        if tip["win"] is not None:
            tip["win"].destroy()
            tip["win"] = None
    widget.unbind("<Enter>")
    widget.unbind("<Leave>")
    widget.bind("<Enter>", _show)
    widget.bind("<Leave>", _hide)

# ── Country name lookup (offline, from assets/countries_*.json) ───────────
_country_names = None
def get_country_name(code: str) -> str:
    """Return the country name for an ISO alpha-2 code, based on TRANSLATE setting.
    Returns the code itself if not found."""
    global _country_names
    if not code:
        return ""
    code = code.upper()
    if _country_names is None:
        lang_file = "countries_es.json" if TRANSLATE_ENABLED else "countries_en.json"
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", lang_file)
        try:
            import json as _json
            with open(path, "r", encoding="utf-8") as f:
                _country_names = _json.load(f)
        except Exception:
            _country_names = {}
    return _country_names.get(code, code)


# ── Per-ticker notes stored in assets/notes/{TICKER}.txt ──────────────────
_NOTES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "notes")

def read_ticker_notes(ticker: str) -> str:
    """Return the text stored in assets/notes/{TICKER}.txt, or empty string if none."""
    if not ticker:
        return ""
    path = os.path.join(_NOTES_DIR, f"{ticker.upper()}.txt")
    if not os.path.exists(path):
        return ""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return ""

def write_ticker_notes(ticker: str, content: str):
    """Write text to assets/notes/{TICKER}.txt. Creates the folder if missing."""
    if not ticker:
        return
    try:
        os.makedirs(_NOTES_DIR, exist_ok=True)
        path = os.path.join(_NOTES_DIR, f"{ticker.upper()}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception:
        pass

DILUTION_API_URL = "https://eapi.askedgar.io/enterprise/v1/dilution-rating"
DILUTION_API_KEY = ASKEDGAR_API_KEY
NEWS_API_URL = "https://eapi.askedgar.io/enterprise/v1/news"
NEWS_API_KEY = ASKEDGAR_API_KEY
DILDATA_API_URL = "https://eapi.askedgar.io/enterprise/v1/dilution-data"
DILDATA_API_KEY = ASKEDGAR_API_KEY
SCREENER_API_URL = "https://eapi.askedgar.io/enterprise/v1/screener"
SCREENER_API_KEY = ASKEDGAR_API_KEY
CHART_ANALYSIS_URL = "https://eapi.askedgar.io/v1/ai-chart-analysis"
CHART_ANALYSIS_KEY = ASKEDGAR_API_KEY
GAP_STATS_URL = "https://eapi.askedgar.io/v1/gap-stats"
GAP_STATS_KEY = ASKEDGAR_API_KEY
OFFERINGS_API_URL = "https://eapi.askedgar.io/v1/offerings"
OFFERINGS_API_KEY = ASKEDGAR_API_KEY
OWNERSHIP_API_URL = "https://eapi.askedgar.io/v1/ownership"
OWNERSHIP_API_KEY = ASKEDGAR_API_KEY
PUMP_DUMP_API_URL = "https://eapi.askedgar.io/v1/pump-and-dump-tracker"
RESEARCH_FULL_API_URL = "https://eapi.askedgar.io/v1/research-reports"
RESEARCH_SHORT_API_URL = "https://eapi.askedgar.io/v1/research-reports-short"
RESEARCH_TLDR_API_URL = "https://eapi.askedgar.io/v1/research-reports-tldr"
REVERSE_SPLITS_API_URL = "https://eapi.askedgar.io/v1/reverse-splits"
SPLIT_STATUS_API_URL = "https://eapi.askedgar.io/v1/split-status"
NASDAQ_COMPLIANCE_API_URL = "https://eapi.askedgar.io/v1/nasdaq-compliance"
MARKET_STRENGTH_API_URL = "https://eapi.askedgar.io/v1/market-strength"
POLL_INTERVAL = 1.0

# TradingView real-time data (session cookie for live prices)
TRADINGVIEW_SESSION_ID = os.environ.get("TRADINGVIEW_SESSION_ID", "")
try:
    GAINERS_REFRESH_SECS = int(os.environ.get("GAINERS_REFRESH_SECS", "60"))
    if GAINERS_REFRESH_SECS < 10:
        GAINERS_REFRESH_SECS = 60
except (ValueError, TypeError):
    GAINERS_REFRESH_SECS = 60
MIN_GAINER_PCT = 15  # minimum % change to show in gainers list
try:
    HIGHLIGHT_NEWS_DAYS = int(os.environ.get("HIGHLIGHT_NEWS_DAYS", "2"))
    if HIGHLIGHT_NEWS_DAYS < 0:
        HIGHLIGHT_NEWS_DAYS = 2
except (ValueError, TypeError):
    HIGHLIGHT_NEWS_DAYS = 2
MIN_GAINER_VOLUME = 500_000  # minimum premarket volume to show in gainers list
FEED_VISIBLE_COUNT = 3  # news/8-K/6-K shown by default in Feed (rest hidden behind "Show more")
JMT415_VISIBLE_COUNT = 5  # JMT415 notes shown by default (rest hidden behind "Show more")

# ── API cache (reduces redundant Ask Edgar calls) ─────────────────────────
CACHE_TTL_SECS = float("inf")  # permanent by default (clear manually via reload button)
NEWS_CACHE_TTL_SECS = 1800  # 30 minutes — news is the only endpoint that expires
_api_cache = {}  # key -> (timestamp, data)

def _cached_fetch(cache_key: str, fetch_fn, ttl: int = CACHE_TTL_SECS):
    """Return cached result if fresh, otherwise call fetch_fn and cache it.
    Does NOT cache None results (transient errors should be retried)."""
    now = time.time()
    if cache_key in _api_cache:
        ts, data = _api_cache[cache_key]
        if now - ts < ttl:
            return data
    data = fetch_fn()
    if data is not None:
        _api_cache[cache_key] = (now, data)
    return data

def _clear_ticker_cache(ticker: str):
    """Remove all cached Ask Edgar API responses for a ticker (for force-reload)."""
    suffix = f":{ticker}"
    keys = [k for k in _api_cache if k.endswith(suffix)]
    for k in keys:
        del _api_cache[k]


# ── API call log (session-scoped; see .env API_LOG=TRUE) ──────────────────
API_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "api.log")
_log_lock = threading.Lock()
_session_cost_usd = 0.0
_session_calls = 0
_session_errors = 0
_session_start_ts = None
_session_by_ticker: dict[str, float] = {}
_session_by_endpoint: dict[str, float] = {}
_session_cost_label = None  # tk.Label, set during _build_ui when API_LOG=TRUE
_ticker_cost_label = None   # tk.Label, ditto — shows cost spent on current ticker
_current_log_ticker = None  # mirror of self.current_ticker, for background log callback

# ── AskEdgar daily call counter (always active, regardless of API_LOG) ────
_edgar_daily_calls = 0
_edgar_call_day = None
_edgar_call_lock = threading.Lock()

def _track_edgar_call():
    """Incrementa el contador diario de llamadas a AskEdgar y actualiza la UI."""
    global _edgar_daily_calls, _edgar_call_day
    from datetime import date as _date_cls
    with _edgar_call_lock:
        today = _date_cls.today()
        if _edgar_call_day != today:
            _edgar_daily_calls = 0
            _edgar_call_day = today
        _edgar_daily_calls += 1
        calls = _edgar_daily_calls
    lbl = _session_cost_label
    if lbl is not None:
        cost = _session_cost_usd if API_LOG_ENABLED else 0.0
        try:
            lbl.after(0, lambda c=cost, k=calls: lbl.config(
                text=f" \U0001f4ca AskEdgar  ${c:.4f}  ({k}calls) "
            ))
        except Exception:
            pass

# ── Claude API cost tracking ──────────────────────────────────────────────
# Precios oficiales Anthropic (USD por millón de tokens, abril 2026)
# claude-sonnet-4-20250514: $3.00 input / $15.00 output
# claude-haiku-4-5-20251001: $0.80 input / $4.00 output
_CLAUDE_PRICE = {
    "claude-sonnet-4-20250514": {"input": 3.00 / 1_000_000, "output": 15.00 / 1_000_000},
    "claude-haiku-4-5-20251001": {"input": 0.80 / 1_000_000, "output": 4.00 / 1_000_000},
}
_claude_daily_cost_usd = 0.0   # acumulado del día (se resetea al cambiar de día)
_claude_daily_calls = 0
_claude_cost_day = None        # fecha (date) del último reset
_claude_cost_lock = threading.Lock()
_claude_cost_label = None      # tk.Label del recuadro de coste en la UI

def _track_claude_cost(model: str, input_tokens: int, output_tokens: int):
    """Registra el coste de una llamada a la API de Claude y actualiza la UI."""
    global _claude_daily_cost_usd, _claude_daily_calls, _claude_cost_day
    from datetime import date as _date_cls
    prices = _CLAUDE_PRICE.get(model, {"input": 3.00 / 1_000_000, "output": 15.00 / 1_000_000})
    cost = input_tokens * prices["input"] + output_tokens * prices["output"]
    with _claude_cost_lock:
        today = _date_cls.today()
        if _claude_cost_day != today:
            # Nuevo día → reset contador
            _claude_daily_cost_usd = 0.0
            _claude_daily_calls = 0
            _claude_cost_day = today
        _claude_daily_cost_usd += cost
        _claude_daily_calls += 1
        daily = _claude_daily_cost_usd
        calls = _claude_daily_calls
    lbl = _claude_cost_label
    if lbl is not None:
        try:
            lbl.after(0, lambda: lbl.config(
                text=f" ⚡ Claude  ${daily:.4f}  ({calls}calls) "
            ))
        except Exception:
            pass
    return cost

# Per-endpoint price table ($/GB), fetched from /estimate on 2026-04-22.
_ASKEDGAR_PRICE_PER_GB = {
    "dilution-rating": 5600, "float-outstanding": 35000, "nasdaq-compliance": 18750,
    "offerings": 6250, "dilution-data": 12500, "historical-float": 7750,
    "historical-float-pro": 15000, "gap-stats": 4000, "news": 500, "news-basic": 50,
    "registrations": 36000, "ownership": 800, "pump-and-dump-tracker": 15000,
    "screener": 10000, "ai-chart-analysis": 2500, "research-reports": 3250,
    "research-reports-short": 4250, "research-reports-tldr": 5750,
    "reverse-splits": 11500, "split-status": 22500, "market-strength": 1000,
    "filing-titles": 20000,
}

def _fmt_bytes(n: int) -> str:
    if n < 1024:
        return f"{n}B"
    if n < 1024 * 1024:
        return f"{n/1024:.1f}KB"
    return f"{n/1024/1024:.1f}MB"

def _log_api_call(ticker: str, endpoint: str, bytes_returned: int,
                  cost_usd: float, duration_ms: int,
                  estimated: bool = False, error: str | None = None):
    """Append one line to api.log and update session counters.
    The call counter is always tracked regardless of API_LOG setting.

    estimated=True marks the cost as locally estimated (not reported by the API),
    which prints a leading '~' before the dollar amount.
    """
    # Always track the call count (regardless of API_LOG_ENABLED)
    _track_edgar_call()

    if not API_LOG_ENABLED:
        return
    global _session_cost_usd, _session_calls, _session_errors, _session_start_ts
    with _log_lock:
        if _session_start_ts is None:
            _session_start_ts = datetime.now()
            try:
                with open(API_LOG_PATH, "a", encoding="utf-8") as f:
                    f.write(f"\n=== Session started: {_session_start_ts.strftime('%Y-%m-%d %H:%M:%S')} ===\n")
            except Exception:
                pass
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if error:
            line = f"{ts}  {ticker:<6} {endpoint:<22}  ERR  —        —             {duration_ms:>5}ms  \"{error}\"\n"
            _session_errors += 1
        else:
            size_str = _fmt_bytes(bytes_returned)
            cost_str = f"{'~$' if estimated else '$'}{cost_usd:.6f}"
            line = f"{ts}  {ticker:<6} {endpoint:<22}  OK   {size_str:<7}  {cost_str:<13} {duration_ms:>5}ms\n"
            _session_cost_usd += cost_usd
            _session_by_ticker[ticker] = _session_by_ticker.get(ticker, 0.0) + cost_usd
            _session_by_endpoint[endpoint] = _session_by_endpoint.get(endpoint, 0.0) + cost_usd
        _session_calls += 1
        try:
            with open(API_LOG_PATH, "a", encoding="utf-8") as f:
                f.write(line)
        except Exception:
            pass
    lbl = _session_cost_label
    if lbl is not None:
        try:
            cost_snap = _session_cost_usd
            calls_snap = _edgar_daily_calls
            lbl.after(0, lambda c=cost_snap, k=calls_snap: lbl.config(
                text=f" 📊 AskEdgar  ${c:.4f}  ({k}calls) "
            ))
        except Exception:
            pass
    tlbl = _ticker_cost_label
    if tlbl is not None and ticker == _current_log_ticker:
        try:
            tc = _session_by_ticker.get(ticker, 0.0)
            tlbl.after(0, lambda: tlbl.config(text=f" ~${tc:.4f} "))
        except Exception:
            pass

def _write_session_api_summary():
    """Append a summary block to api.log when the app closes."""
    if not API_LOG_ENABLED or _session_start_ts is None:
        return
    end_ts = datetime.now()
    by_ticker = sorted(_session_by_ticker.items(), key=lambda kv: kv[1], reverse=True)[:5]
    by_endpoint = sorted(_session_by_endpoint.items(), key=lambda kv: kv[1], reverse=True)[:5]
    lines = [
        f"=== Session ended:   {end_ts.strftime('%Y-%m-%d %H:%M:%S')} "
        f"(duration: {str(end_ts - _session_start_ts).split('.')[0]}) ===",
        f"Total API calls: {_session_calls} ({_session_errors} errors)",
        f"Total cost: ${_session_cost_usd:.4f}",
    ]
    if by_ticker:
        lines.append("Top tickers: " + ", ".join(f"{t} ${c:.4f}" for t, c in by_ticker))
    if by_endpoint:
        lines.append("Top endpoints: " + ", ".join(f"{e} ${c:.4f}" for e, c in by_endpoint))
    lines.append("")
    try:
        with _log_lock, open(API_LOG_PATH, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
    except Exception:
        pass

if not ASKEDGAR_API_KEY:
    print("ERROR: Missing API key. Copy .env.example to .env and fill in your key.")
    print("  ASKEDGAR_API_KEY - request trial at https://www.askedgar.io/api-trial")

# Ticker filter: 2-4 uppercase letters, no periods or special chars
TICKER_RE = re.compile(r'^[A-Z]{2,4}$')

# ── Visual Style ────────────────────────────────────────────────────────────
BG = "#0D1014"
BG_CARD = "#151A20"
BG_ROW = "#1B2128"
BG_ROW_ALT = "#181D24"
BG_SELECTED = "#1A2A3A"
BORDER = "#232A33"
BORDER_INNER = "#20262E"
BORDER_ACCENT = "#63D3FF"
FG = "#E6EAF0"
FG_DIM = "#8B949E"
FG_INFO = "#B7C0CC"
ACCENT = "#63D3FF"
GREEN = "#4CAF50"
RED = "#FF4444"

RISK_BG = {
    "High": "#A93232",
    "Medium": "#B96A16",
    "Low": "#2F7D57",
    "N/A": "#4A525C",
}

# Chart history rating: API color -> (label, badge color)
HISTORY_MAP = {
    "green":  ("Strong", "#2F7D57"),
    "yellow": ("Semi-Strong", "#B9A816"),
    "orange": ("Mixed",      "#B96A16"),
    "red":    ("Fader",  "#A93232"),
}

# Fonts
FONT_UI = ("Segoe UI", 10)
FONT_UI_BOLD = ("Segoe UI Semibold", 10)
FONT_HEADER = ("Segoe UI Semibold", 13)
FONT_TICKER = ("Segoe UI Semibold", 24)
FONT_MONO = ("Consolas", 9)
FONT_MONO_BOLD = ("Consolas", 9, "bold")
FONT_GAINER_TICKER = ("Segoe UI Semibold", 12)
FONT_GAINER_PCT = ("Consolas", 11, "bold")
FONT_GAINER_DETAIL = ("Consolas", 8)

LEFT_PANEL_WIDTH = 260
MONTAGE_PANEL_WIDTH = 250


def risk_bg(level: str) -> str:
    return RISK_BG.get(level, "#555555")


def fmt_millions(val) -> str:
    if val is None:
        return "N/A"
    m = val / 1_000_000
    if m >= 1:
        return f"{m:.2f}M"
    return f"{val / 1000:.0f}K"


def fmt_volume(val) -> str:
    """Format volume with K/M suffix."""
    if val is None or val == 0:
        return "0"
    if val >= 1_000_000:
        return f"{val / 1_000_000:.1f}M"
    if val >= 1_000:
        return f"{val / 1_000:.0f}K"
    return str(int(val))


def fmt_price(val) -> str:
    """Format price with appropriate decimal places."""
    if val is None or val == 0:
        return "$0.00"
    if val >= 1:
        return f"${val:.2f}"
    return f"${val:.4f}"


# ── Window Monitor ──────────────────────────────────────────────────────────
_das_debug_printed: set = set()  # track which unrecognized DAS titles we've logged

def find_montage_windows() -> dict[int, str]:
    """Return {hwnd: ticker} for all visible DAS montage and chart windows.

    Supports windows on any monitor (primary or secondary).
    DAS montage title formats observed:
      "AAPL     0 -- 0     Apple Inc"       (classic)
      "AAPL  123.45  +1.23%  Apple Inc"     (with price)
      "AAPL--5 Minute--"                    (chart window)
      "AAPL - Montage"                      (alternative)
    """
    windows = {}

    def enum_callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return

        # DAS montage: "TICKER     0 -- 0     Company Name..."
        # or "TICKER  123.45  ..."  (price in title)
        if re.match(r'^[A-Z]{1,5}\s+[\d\-]', title):
            candidate = title.split()[0]
            if re.match(r'^[A-Z]{1,5}$', candidate):
                windows[hwnd] = candidate
                return

        # DAS chart: "TICKER--5 Minute--"
        if re.match(r'^[A-Z]{1,5}--', title):
            candidate = title.split('--')[0]
            if re.match(r'^[A-Z]{1,5}$', candidate):
                windows[hwnd] = candidate
                return

        # DAS alternative: "TICKER - Montage" or "TICKER - Chart"
        m = re.match(r'^([A-Z]{1,5})\s*[-–]\s*(Montage|Chart|Level|Quote)', title)
        if m:
            windows[hwnd] = m.group(1)
            return

        # Debug: log any window that contains "DAS" or "Montage" in its title
        # so we can identify unrecognized formats
        if ("DAS" in title or "Montage" in title or "montage" in title) and hwnd not in _das_debug_printed:
            _das_debug_printed.add(hwnd)
            print(f"[DEBUG] Unrecognized DAS window title: {repr(title)}")

    win32gui.EnumWindows(enum_callback, None)
    return windows


def find_tos_tickers() -> dict[int, list[str]]:
    """Return {hwnd: [tickers]} for thinkorswim chart windows."""
    windows = {}

    def enum_callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        # "PRSO, MOBX, TURB - Charts - 61612650SCHW Main@thinkorswim [build 1990]"
        if "thinkorswim" in title and " - Charts - " in title:
            ticker_part = title.split(" - Charts - ")[0]
            tickers = [t.strip() for t in ticker_part.split(",") if t.strip()]
            if tickers:
                windows[hwnd] = tickers

    win32gui.EnumWindows(enum_callback, None)
    return windows


# ── Market Data APIs ────────────────────────────────────────────────────────
def _tv_cookies():
    """Build TradingView cookie jar for real-time data."""
    jar = requests.cookies.RequestsCookieJar()
    if TRADINGVIEW_SESSION_ID:
        jar.set("sessionid", TRADINGVIEW_SESSION_ID, domain=".tradingview.com")
    return jar


def fetch_top_gainers() -> list[dict]:
    """Fetch top premarket gainers from TradingView (real-time), enriched with Ask Edgar data."""
    try:
        from tradingview_screener import Query, col
    except ImportError:
        print("tradingview-screener not installed. Run: pip install tradingview-screener")
        return []

    try:
        cookies = _tv_cookies()
        _, df = (Query()
            .select("name", "close", "premarket_change", "premarket_close",
                    "premarket_volume", "volume", "market_cap_basic")
            .where(col("premarket_change") > MIN_GAINER_PCT)
            .order_by("premarket_change", ascending=False)
            .limit(30)
            .get_scanner_data(cookies=cookies if TRADINGVIEW_SESSION_ID else None))
    except Exception as e:
        print(f"TradingView scanner error: {e}")
        return []

    # Convert DataFrame to our internal format
    tickers_data = []
    for _, row in df.iterrows():
        ticker = row.get("name", "")
        if not TICKER_RE.match(ticker):
            continue
        # Extract exchange from "NASDAQ:SKYQ" format
        exchange = ""
        full_ticker = row.get("ticker", "")
        if ":" in full_ticker:
            exchange = full_ticker.split(":")[0]
        pct = row.get("premarket_change") or 0
        tickers_data.append({
            "ticker": ticker,
            "_exchange": exchange,
            "todaysChangePerc": pct,
            "price": row.get("premarket_close") or row.get("close") or 0,
            "volume": int(row.get("premarket_volume") or row.get("volume") or 0),
            "_tv_mcap": row.get("market_cap_basic") or 0,
        })

    # Enrich with Ask Edgar data in parallel
    def enrich(item, include_dilution=False):
        ticker = item["ticker"]
        sdata = fetch_screener_data(ticker)
        if sdata:
            item["_float"] = sdata.get("tradable_float")
            item["_mcap"] = sdata.get("market_cap")
            item["_sector"] = sdata.get("sector", "")
            item["_country"] = sdata.get("country", "")
        if include_dilution:
            ddata = fetch_dilution_data(ticker)
            if ddata:
                item["_risk"] = ddata.get("overall_offering_risk", "")
        # Check for news/filings today (uses cached news data)
        from datetime import datetime
        today = datetime.now().strftime("%Y-%m-%d")
        news_results = _cached_news_results(ticker)
        if news_results is not None:
            for r in news_results:
                ft = r.get("form_type")
                if ft in ("news", "8-K", "6-K"):
                    d = (r.get("created_at") or r.get("filed_at", ""))[:10]
                    if d == today:
                        item["_news_today"] = True
                        break
        return item

    enriched = []
    with ThreadPoolExecutor(max_workers=6) as executor:
        futures = {executor.submit(enrich, item, True): item for i, item in enumerate(tickers_data[:30])}
        for future in futures:
            result = future.result()
            if result is not None:
                enriched.append(result)

    enriched.sort(key=lambda x: x.get("todaysChangePerc", 0), reverse=True)
    return enriched


def fetch_tv_ticker_data(ticker: str) -> dict:
    """Fetch price, premarket change, and volume for a single ticker from TradingView screener."""
    try:
        from tradingview_screener import Query, col
        cookies = _tv_cookies()
        _, df = (Query()
            .select("name", "close", "premarket_change", "premarket_close",
                    "premarket_volume", "volume", "change")
            .where(col("name") == ticker)
            .limit(1)
            .get_scanner_data(cookies=cookies if TRADINGVIEW_SESSION_ID else None))
        if df is not None and not df.empty:
            row = df.iloc[0]
            pct = row.get("premarket_change") or row.get("change") or 0
            price = row.get("premarket_close") or row.get("close") or 0
            volume = int(row.get("premarket_volume") or row.get("volume") or 0)
            return {"todaysChangePerc": pct, "price": price, "volume": volume}
    except Exception as e:
        print(f"TV ticker data error for {ticker}: {e}")
    return {}


# ── Ask Edgar APIs ──────────────────────────────────────────────────────────
def fetch_dilution_data(ticker: str) -> dict | None:
    def _fetch():
        try:
            resp = requests.get(
                DILUTION_API_URL,
                headers={"API-KEY": DILUTION_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "offset": 0, "limit": 10},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success" and data.get("results"):
                print(f"[DILUTION] {ticker}: OK, risk={data['results'][0].get('overall_offering_risk')}")
                return data["results"][0]
            else:
                print(f"[DILUTION] {ticker}: no data, status={resp.status_code}, body={data}")
        except Exception as e:
            print(f"Dilution API error for {ticker}: {e}")
        return None
    return _cached_fetch(f"dilution:{ticker}", _fetch)



def _cached_news_results(ticker: str) -> list[dict] | None:
    """Fetch raw news results for a ticker, cached for 5 minutes."""
    def _fetch():
        try:
            resp = requests.get(
                NEWS_API_URL,
                headers={"API-KEY": NEWS_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "offset": 0, "limit": 100},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success":
                results = data.get("results", [])
                form_types = [r.get("form_type") for r in results]
                print(f"[NEWS] {ticker}: {len(results)} items, types={form_types[:10]}, status={resp.status_code}")
                return results
            else:
                print(f"[NEWS] {ticker}: API error status={resp.status_code}, body={data}")
        except Exception as e:
            print(f"News API error for {ticker}: {e}")
        return None
    return _cached_fetch(f"news:{ticker}", _fetch)


def fetch_news_and_grok(ticker: str) -> tuple[list[dict], str | None, str | None, str | None, list[dict]]:
    """Fetch recent news/8-K/6-K (top 2), latest grok, and all jmt415 notes."""
    headlines = []
    grok_line = None
    grok_date = None
    grok_url = None
    jmt415_notes = []
    results = _cached_news_results(ticker)
    if results:
        for r in results:
            ft = r.get("form_type")
            if ft in ("news", "8-K", "6-K") and len(headlines) < 2:
                headlines.append(r)
            if ft == "grok" and grok_line is None:
                summary = r.get("summary", "")
                for line in summary.split("\n"):
                    line = line.strip().lstrip("-").strip()
                    if line:
                        grok_line = line
                        break
                # created_at includes time, fall back to filed_at
                grok_date = r.get("created_at") or r.get("filed_at", "")
                grok_url = (r.get("url") or r.get("document_url") or
                            f"https://app.askedgar.io/ticker/{ticker}/news")
            if ft == "jmt415" and len(jmt415_notes) < 3:
                jmt415_notes.append(r)
        print(f"[NEWS] {ticker}: feed_headlines={len(headlines)}, grok={'yes' if grok_line else 'no'}, jmt415={len(jmt415_notes)}")
    else:
        print(f"[NEWS] {ticker}: no results returned (None or empty)")
    return headlines, grok_line, grok_date, grok_url, jmt415_notes


def fetch_screener_data(ticker: str) -> dict | None:
    """Fetch screener data (price, float, outstanding, sector, country, mcap) via Ask Edgar screener endpoint."""
    def _fetch():
        try:
            resp = requests.get(
                SCREENER_API_URL,
                headers={"API-KEY": SCREENER_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success" and data.get("results"):
                return data["results"][0]
            else:
                print(f"Screener API [{resp.status_code}] {ticker}: {data}")
        except Exception as e:
            print(f"Screener API error for {ticker}: {e}")
        return None
    return _cached_fetch(f"screener:{ticker}", _fetch)


def fetch_in_play_dilution(ticker: str) -> tuple[list[dict], list[dict], float]:
    """Fetch dilution-data and split into in-play warrants and convertibles.
    Returns (warrants, convertibles, stock_price) filtered by price proximity and registration."""
    def _fetch():
        sdata = fetch_screener_data(ticker)
        price = sdata.get("price") if sdata else None
        if price is None or price <= 0:
            return [], [], 0.0

        max_price = price * 4

        try:
            resp = requests.get(
                DILDATA_API_URL,
                headers={"API-KEY": DILDATA_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "offset": 0, "limit": 40},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") != "success":
                print(f"Dilution-data API [{resp.status_code}] {ticker}: {data}")
                return [], [], price
        except Exception as e:
            print(f"Dilution-data API error for {ticker}: {e}")
            return [], [], price

        warrants = []
        convertibles = []
        from datetime import datetime, timedelta
        six_months_ago = datetime.now() - timedelta(days=180)

        for item in data.get("results", []):
            registered = item.get("registered") or ""
            details_lower = (item.get("details") or "").lower()
            is_warrant = "warrant" in details_lower or "option" in details_lower

            # Skip "Not Registered" items, but override for convertibles filed >6 months ago
            skip_not_registered = "Not Registered" in registered
            if skip_not_registered and not is_warrant:
                filed_at_str = (item.get("filed_at") or "")[:10]
                if filed_at_str:
                    try:
                        if datetime.strptime(filed_at_str, "%Y-%m-%d") < six_months_ago:
                            skip_not_registered = False
                    except ValueError:
                        pass
            if skip_not_registered:
                continue

            if is_warrant and item.get("warrants_exercise_price"):
                if item["warrants_exercise_price"] <= max_price:
                    remaining = item.get("warrants_remaining", 0) or 0
                    if remaining > 0:
                        warrants.append(item)
            elif not is_warrant and item.get("conversion_price"):
                if item["conversion_price"] <= max_price:
                    remaining = item.get("underlying_shares_remaining", 0) or 0
                    if remaining > 0:
                        convertibles.append(item)

        return warrants, convertibles, price
    return _cached_fetch(f"inplay:{ticker}", _fetch) or ([], [], 0.0)


def fetch_gap_stats(ticker: str) -> list[dict]:
    """Fetch gap-up stats for a ticker. Returns list of gap entries (date descending)."""
    def _fetch():
        try:
            resp = requests.get(
                GAP_STATS_URL,
                headers={"API-KEY": GAP_STATS_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "page": 1, "limit": 100},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success":
                return data.get("results", [])
            else:
                print(f"Gap stats API [{resp.status_code}] {ticker}: {data}")
        except Exception as e:
            print(f"Gap stats API error for {ticker}: {e}")
        return []
    return _cached_fetch(f"gapstats:{ticker}", _fetch) or []


def fetch_offerings(ticker: str) -> list[dict]:
    """Fetch recent offerings for the ticker (up to 5)."""
    def _fetch():
        try:
            resp = requests.get(
                OFFERINGS_API_URL,
                headers={"API-KEY": OFFERINGS_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "limit": 5},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success":
                return data.get("results", [])
            else:
                print(f"Offerings API [{resp.status_code}] {ticker}: {data}")
        except Exception as e:
            print(f"Offerings API error for {ticker}: {e}")
        return []
    return _cached_fetch(f"offerings:{ticker}", _fetch) or []


def fetch_ownership(ticker: str) -> dict | None:
    """Fetch ownership data – returns the latest reported_date group, or None."""
    def _fetch():
        try:
            resp = requests.get(
                OWNERSHIP_API_URL,
                headers={"API-KEY": OWNERSHIP_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "limit": 100},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success" and data.get("results"):
                return data["results"][0]  # latest reported_date group
            else:
                print(f"Ownership API [{resp.status_code}] {ticker}: {data}")
        except Exception as e:
            print(f"Ownership API error for {ticker}: {e}")
        return None
    return _cached_fetch(f"ownership:{ticker}", _fetch)


def fetch_chart_analysis(ticker: str) -> dict | None:
    """Fetch chart analysis (history rating). Returns first result dict or None."""
    def _fetch():
        try:
            resp = requests.get(
                CHART_ANALYSIS_URL,
                headers={"API-KEY": CHART_ANALYSIS_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "limit": 1},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success" and data.get("results"):
                return data["results"][0]
            else:
                print(f"Chart API [{resp.status_code}] {ticker}: {data}")
        except Exception as e:
            print(f"Chart API error for {ticker}: {e}")
        return None
    return _cached_fetch(f"chart:{ticker}", _fetch)


def capture_das_chart_window(ticker: str) -> bytes | None:
    """Capture a screenshot of the DAS Trader chart window for the given ticker.
    Returns PNG bytes or None if no chart window found."""
    try:
        import win32gui
        import win32con
        import win32ui
        import ctypes
        from PIL import Image
        import io

        target_hwnd = None

        def enum_cb(hwnd, _):
            nonlocal target_hwnd
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = win32gui.GetWindowText(hwnd)
            # Match DAS chart windows: "TICKER--5 Minute--" or "TICKER - Chart"
            if re.match(rf'^{re.escape(ticker)}--', title):
                target_hwnd = hwnd
            elif re.match(rf'^{re.escape(ticker)}\s*[-–]\s*Chart', title, re.IGNORECASE):
                target_hwnd = hwnd
            # Also accept any DAS montage window for this ticker
            elif re.match(rf'^{re.escape(ticker)}\s+[\d\-]', title):
                if target_hwnd is None:
                    target_hwnd = hwnd

        win32gui.EnumWindows(enum_cb, None)

        if not target_hwnd:
            print(f"[Overhead] No DAS window found for {ticker}")
            return None

        # Get window rect
        left, top, right, bottom = win32gui.GetWindowRect(target_hwnd)
        width = right - left
        height = bottom - top
        if width <= 0 or height <= 0:
            return None

        # Capture using PIL ImageGrab (simpler, works for on-screen windows)
        from PIL import ImageGrab
        img = ImageGrab.grab(bbox=(left, top, right, bottom))
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    except Exception as e:
        print(f"[Overhead] Screenshot error: {e}")
        return None


def analyze_overhead_with_claude(ticker: str, image_bytes: bytes) -> bool:
    """Send chart screenshot to Claude Vision and ask if there is overhead resistance.
    Returns True if overhead detected, False otherwise."""
    if not CLAUDE_API_KEY or not image_bytes:
        return False
    try:
        import base64
        img_b64 = base64.b64encode(image_bytes).decode()
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": CLAUDE_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 100,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/png",
                                    "data": img_b64,
                                },
                            },
                            {
                                "type": "text",
                                "text": (
                                    f"This is a stock chart for {ticker}. "
                                    "Does this chart have significant overhead resistance "
                                    "(prior price levels, supply zones, or gaps above the current price "
                                    "that could act as resistance)? "
                                    "Reply with only YES or NO."
                                ),
                            },
                        ],
                    }
                ],
            },
            timeout=15,
        )
        data = resp.json()
        answer = ""
        for block in data.get("content", []):
            if block.get("type") == "text":
                answer = block["text"].strip().upper()
                break
        # Track Claude API cost
        usage = data.get("usage", {})
        _track_claude_cost("claude-sonnet-4-20250514",
                           usage.get("input_tokens", 0),
                           usage.get("output_tokens", 0))
        print(f"[Overhead] Claude says for {ticker}: {answer}")
        return answer.startswith("YES")
    except Exception as e:
        print(f"[Overhead] Claude Vision error: {e}")
        return False


def fetch_company_name(cik: str) -> str:
    """Fetch company name from SEC EDGAR by CIK. Returns empty string on failure."""
    if not cik or not SHOW_COMPANY_NAME:
        return ""
    padded = cik.zfill(10)
    def _fetch():
        try:
            resp = requests.get(
                f"https://data.sec.gov/submissions/CIK{padded}.json",
                headers={"User-Agent": "DilutionMonitor/1.0"},
                timeout=5,
            )
            if resp.status_code == 200:
                return resp.json().get("name", "")
        except Exception:
            pass
        return ""
    return _cached_fetch(f"company:{cik}", _fetch) or ""


def fetch_pump_dump(ticker: str) -> list[dict]:
    """Fetch pump & dump tracker data for a ticker."""
    def _fetch():
        try:
            resp = requests.get(
                PUMP_DUMP_API_URL,
                headers={"API-KEY": ASKEDGAR_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "limit": 1},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success":
                return data.get("results", [])
        except Exception as e:
            print(f"Pump&Dump API error for {ticker}: {e}")
        return []
    return _cached_fetch(f"pumpdump:{ticker}", _fetch) or []


def fetch_research_full(ticker: str) -> list[dict]:
    """Fetch full research report for a ticker."""
    def _fetch():
        try:
            resp = requests.get(
                RESEARCH_FULL_API_URL,
                headers={"API-KEY": ASKEDGAR_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "limit": 1},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success":
                return data.get("results", [])
        except Exception as e:
            print(f"Research Full API error for {ticker}: {e}")
        return []
    return _cached_fetch(f"research_full:{ticker}", _fetch) or []


def fetch_research_short(ticker: str) -> list[dict]:
    """Fetch short research report for a ticker."""
    def _fetch():
        try:
            resp = requests.get(
                RESEARCH_SHORT_API_URL,
                headers={"API-KEY": ASKEDGAR_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "limit": 1},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success":
                return data.get("results", [])
        except Exception as e:
            print(f"Research Short API error for {ticker}: {e}")
        return []
    return _cached_fetch(f"research_short:{ticker}", _fetch) or []


def fetch_research_tldr(ticker: str) -> list[dict]:
    """Fetch TLDR research report for a ticker."""
    def _fetch():
        try:
            resp = requests.get(
                RESEARCH_TLDR_API_URL,
                headers={"API-KEY": ASKEDGAR_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "limit": 1},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success":
                return data.get("results", [])
        except Exception as e:
            print(f"Research TLDR API error for {ticker}: {e}")
        return []
    return _cached_fetch(f"research_tldr:{ticker}", _fetch) or []


def fetch_reverse_splits(ticker: str) -> list[dict]:
    """Fetch reverse split history for a ticker."""
    def _fetch():
        try:
            resp = requests.get(
                REVERSE_SPLITS_API_URL,
                headers={"API-KEY": ASKEDGAR_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "limit": 10},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success":
                return data.get("results", [])
        except Exception as e:
            print(f"Reverse Splits API error for {ticker}: {e}")
        return []
    return _cached_fetch(f"rsplits:{ticker}", _fetch) or []


def fetch_split_status(ticker: str) -> list[dict]:
    """Fetch pending/announced split status for a ticker."""
    def _fetch():
        try:
            resp = requests.get(
                SPLIT_STATUS_API_URL,
                headers={"API-KEY": ASKEDGAR_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "limit": 5},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success":
                return data.get("results", [])
        except Exception as e:
            print(f"Split Status API error for {ticker}: {e}")
        return []
    return _cached_fetch(f"splitstatus:{ticker}", _fetch) or []


def fetch_nasdaq_compliance(ticker: str) -> list[dict]:
    """Fetch NASDAQ compliance status for a ticker."""
    def _fetch():
        try:
            resp = requests.get(
                NASDAQ_COMPLIANCE_API_URL,
                headers={"API-KEY": ASKEDGAR_API_KEY, "Content-Type": "application/json"},
                params={"ticker": ticker, "limit": 5},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "success":
                return data.get("results", [])
        except Exception as e:
            print(f"NASDAQ Compliance API error for {ticker}: {e}")
        return []
    return _cached_fetch(f"compliance:{ticker}", _fetch) or []


def extract_headline(item: dict) -> str:
    if item.get("title"):
        return item["title"]
    summary = item.get("summary", "")
    if summary.startswith("HEADLINE:"):
        return summary.split("HEADLINE:")[1].split("\n")[0].strip()
    return f"{item.get('form_type', '')} Filing"


def fetch_claude_signal(ticker: str, company_name: str, grok_text: str, news_headlines: list[str]) -> dict | None:
    """Call Claude API to get LONG/SHORT signal with summary. Returns dict with signal, summary or None."""
    if not CLAUDE_API_KEY or not grok_text:
        return None
    try:
        headlines_text = "\n".join(f"- {h}" for h in news_headlines) if news_headlines else "No hay noticias adicionales."
        prompt = f"""Eres un analista experto en day trading de small caps en el mercado americano.

Analiza la siguiente informacion sobre el ticker ${ticker} ({company_name}) y determina si es una oportunidad LONG o SHORT para hoy.

RESUMEN GROK:
{grok_text}

NOTICIAS RECIENTES:
{headlines_text}

Responde UNICAMENTE en este formato JSON sin texto adicional ni markdown:
{{
  "signal": "LONG o SHORT",
  "summary": "Resumen breve en espanol de 1-2 frases del contexto",
  "catalysts": ["catalizador 1", "catalizador 2", "catalizador 3"],
  "action": "Accion recomendada: cuando entrar stop y target aproximado",
  "red_flags": ["red flag 1", "red flag 2", "red flag 3"]
}}"""

        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": CLAUDE_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": "claude-haiku-4-5-20251001",
                "max_tokens": 500,
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=15,
        )
        data = resp.json()
        text = data.get("content", [{}])[0].get("text", "").strip()
        # Track Claude API cost
        usage = data.get("usage", {})
        _track_claude_cost("claude-haiku-4-5-20251001",
                           usage.get("input_tokens", 0),
                           usage.get("output_tokens", 0))
        import json as _json
        text = text.replace("```json", "").replace("```", "").strip()
        result = _json.loads(text)
        signal = result.get("signal", "").upper()
        summary = result.get("summary", "")
        if signal in ("LONG", "SHORT") and summary:
            return {
                "signal": signal,
                "summary": summary,
                "catalysts": result.get("catalysts", []),
                "action": result.get("action", ""),
                "red_flags": result.get("red_flags", []),
            }
    except Exception as e:
        print(f"Claude signal error for {ticker}: {e}")
    return None


# ── Overlay UI ──────────────────────────────────────────────────────────────
class DilutionOverlay:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DILU-PIRATA")
        self.root.attributes("-topmost", True)
        self.root.attributes("-toolwindow", False)
        self.root.configure(bg=BG)
        self.root.resizable(True, True)
        self.root.geometry(self._load_position())
        self.root.minsize(830, 400)

        self._drag_data = {"x": 0, "y": 0}
        self.current_ticker = None
        self._known_windows: dict[int, str] = {}   # DAS: hwnd -> ticker
        self._known_tos: dict[int, list[str]] = {}  # ToS: hwnd -> [tickers]
        self._gainers_data: list[dict] = []
        self._selected_gainer: str | None = None
        self._prev_gainer_tickers: set[str] = set()   # for new-ticker alert
        self._new_gainer_tickers: set[str] = set()    # tickers that appeared this refresh
        self._ticker_history: list[dict] = []          # last 15 tickers viewed
        self._chart_ticker_file = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), ".chart_ticker"
        )
        # Claude history: list of dicts with signal, summary, catalysts, action, red_flags
        self._claude_history: list[dict] = []
        # Store current ticker's data for on-demand ANALIZAR button
        self._current_data_for_claude: dict = {}
        # Session start time (for session summary on close)
        from datetime import datetime as _dt_init
        self._session_start = _dt_init.now()
        self._build_ui()
        self._start_monitor()
        self._schedule_gainers_refresh()

    def _load_position(self) -> str:
        """Load saved window geometry from position.cfg. Returns default if not found."""
        try:
            cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "position.cfg")
            if os.path.exists(cfg_path):
                with open(cfg_path, "r") as f:
                    geo = f.read().strip()
                if geo:
                    return geo
        except Exception:
            pass
        return "780x700+50+50"

    def _save_position(self):
        """Save current window geometry to position.cfg, and chart position if open."""
        import json
        try:
            self.root.update_idletasks()
            geo = self.root.geometry()
            cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "position.cfg")
            with open(cfg_path, "w") as f:
                f.write(geo)
            # Save chart window position via shared file
            # Write a request for the chart process to save its position
            chart_save_file = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), ".chart_save_position"
            )
            with open(chart_save_file, "w") as f:
                f.write("save")
            # Show confirmation toast
            toast = tk.Toplevel(self.root)
            toast.overrideredirect(True)
            toast.attributes("-topmost", True)
            x = self.root.winfo_x() + self.root.winfo_width() // 2 - 120
            y = self.root.winfo_y() + 50
            toast.geometry(f"+{x}+{y}")
            tk.Label(toast, text=" Position saved! ", fg="white", bg="#2F7D57",
                     font=FONT_UI_BOLD, padx=16, pady=8).pack()
            toast.after(1500, toast.destroy)
        except Exception:
            pass

    def _build_ui(self):
        # ── Search bar (top, full width) ──
        search_frame = tk.Frame(self.root, bg=BG_CARD,
                                highlightbackground=BORDER, highlightthickness=1)
        search_frame.pack(fill="x", padx=8, pady=(8, 0))

        search_inner = tk.Frame(search_frame, bg=BG_CARD, padx=10, pady=8)
        search_inner.pack(fill="x")
        search_inner.bind("<Button-1>", self._start_drag)
        search_inner.bind("<B1-Motion>", self._on_drag)

        tk.Label(search_inner, text="TICKER:", fg=FG_DIM, bg=BG_CARD,
                 font=FONT_UI_BOLD).pack(side="left", padx=(0, 6))

        self.search_entry = tk.Entry(
            search_inner, bg=BG_ROW, fg=FG, insertbackground=FG,
            font=FONT_UI_BOLD, width=10, relief="flat",
            highlightbackground=BORDER, highlightthickness=1,
        )
        self.search_entry.pack(side="left", padx=(0, 6), ipady=3)
        self.search_entry.bind("<Return>", self._on_search)

        go_btn = tk.Label(
            search_inner, text="  GO  ", fg=BG, bg=ACCENT,
            font=FONT_UI_BOLD, padx=8, pady=2, cursor="hand2",
        )
        go_btn.pack(side="left")
        go_btn.bind("<Button-1>", self._on_search)

        ms_btn = tk.Label(
            search_inner, text=" Market Strength ", fg=BG, bg="#B96A16",
            font=FONT_UI_BOLD, padx=8, pady=2, cursor="hand2",
        )
        ms_btn.pack(side="left", padx=(6, 0))
        ms_btn.bind("<Button-1>", lambda e: self._open_market_strength())

        chart_btn = tk.Label(
            search_inner, text=" Open Chart ", fg=BG, bg="#1F8FB3",
            font=FONT_UI_BOLD, padx=8, pady=2, cursor="hand2",
        )
        chart_btn.pack(side="left", padx=(6, 0))
        chart_btn.bind("<Button-1>", lambda e: self._open_chart_window())

        save_btn = tk.Label(
            search_inner, text=" Save Position ", fg=BG, bg="#2F7D57",
            font=FONT_UI_BOLD, padx=8, pady=2, cursor="hand2",
        )
        save_btn.pack(side="left", padx=(6, 0))
        save_btn.bind("<Button-1>", lambda e: self._save_position())

        # ── HISTORIAL button ──
        historial_btn = tk.Label(
            search_inner, text=" HISTORIAL ", fg="white", bg="#5A3A7D",
            font=FONT_UI_BOLD, padx=8, pady=2, cursor="hand2",
        )
        historial_btn.pack(side="left", padx=(6, 0))
        historial_btn.bind("<Button-1>", lambda e: self._open_historial())

        # ── RECIENTES button ──
        recientes_btn = tk.Label(
            search_inner, text=" RECIENTES ", fg="white", bg="#1F5A7D",
            font=FONT_UI_BOLD, padx=8, pady=2, cursor="hand2",
        )
        recientes_btn.pack(side="left", padx=(6, 0))
        recientes_btn.bind("<Button-1>", lambda e: self._open_recientes())

        # ── ANALIZAR button ──
        self._analizar_btn = tk.Label(
            search_inner, text=" ⚡ ANALIZAR ", fg="black", bg="#FFD600",
            font=FONT_UI_BOLD, padx=8, pady=2, cursor="hand2",
        )
        self._analizar_btn.pack(side="left", padx=(6, 0))
        self._analizar_btn.bind("<Button-1>", lambda e: self._on_analizar_click())

        # ── Cost box: AskEdgar + Claude (hoy) ──
        cost_frame = tk.Frame(search_inner, bg="#0D1A0D",
                              highlightbackground="#2A3D2A", highlightthickness=1)
        cost_frame.pack(side="left", padx=(10, 0))

        # AskEdgar cost label
        global _session_cost_label
        _session_cost_label = tk.Label(
            cost_frame,
            text=" 📊 AskEdgar  $0.0000  (0calls) ",
            fg="#4CAF50", bg="#0D1A0D",
            font=FONT_MONO, padx=4, pady=2,
        )
        _session_cost_label.pack(side="left")

        # Separator
        tk.Label(cost_frame, text="|", fg="#2A3D2A", bg="#0D1A0D",
                 font=FONT_MONO).pack(side="left")

        # Claude cost label
        global _claude_cost_label
        _claude_cost_label = tk.Label(
            cost_frame,
            text=" ⚡ Claude  $0.0000  (0calls) ",
            fg="#63D3FF", bg="#0D1A0D",
            font=FONT_MONO, padx=4, pady=2,
        )
        _claude_cost_label.pack(side="left")

        title_lbl = tk.Label(search_inner, text="DILU-PIRATA",
                             fg=FG_DIM, bg=BG_CARD, font=FONT_UI)
        title_lbl.pack(side="right")
        title_lbl.bind("<Button-1>", self._start_drag)
        title_lbl.bind("<B1-Motion>", self._on_drag)

        # ── Main body (left + right) ──
        main_body = tk.Frame(self.root, bg=BG)
        main_body.pack(fill="both", expand=True)

        # ── Left panel (gainers) ──
        left_panel = tk.Frame(main_body, bg=BG, width=LEFT_PANEL_WIDTH)
        left_panel.pack(side="left", fill="y", padx=(8, 0), pady=(6, 8))
        left_panel.pack_propagate(False)

        # Gainers header
        gh_frame = tk.Frame(left_panel, bg=BG_CARD,
                            highlightbackground=BORDER, highlightthickness=1)
        gh_frame.pack(fill="x")

        gh_inner = tk.Frame(gh_frame, bg=BG_CARD, padx=10, pady=8)
        gh_inner.pack(fill="x")

        tk.Label(gh_inner, text="TOP GAINERS", fg=ACCENT, bg=BG_CARD,
                 font=FONT_HEADER).pack(side="left")

        self._gainers_status = tk.Label(gh_inner, text="", fg=FG_DIM, bg=BG_CARD,
                                        font=FONT_MONO)
        self._gainers_status.pack(side="right")

        refresh_btn = tk.Label(gh_inner, text=" \u21bb ", fg=ACCENT, bg=BG_CARD,
                               font=("Segoe UI", 14), cursor="hand2")
        refresh_btn.pack(side="right", padx=(0, 4))
        refresh_btn.bind("<Button-1>", lambda e: self._trigger_gainers_refresh())

        # Gainers scrollable list
        gainers_container = tk.Frame(left_panel, bg=BG)
        gainers_container.pack(fill="both", expand=True, pady=(2, 0))

        self._gainers_canvas = tk.Canvas(gainers_container, bg=BG,
                                         highlightthickness=0,
                                         width=LEFT_PANEL_WIDTH - 16)
        gainers_sb = tk.Scrollbar(gainers_container, orient="vertical",
                                  command=self._gainers_canvas.yview)
        self._gainers_frame = tk.Frame(self._gainers_canvas, bg=BG)

        self._gainers_frame.bind(
            "<Configure>",
            lambda e: self._gainers_canvas.configure(
                scrollregion=self._gainers_canvas.bbox("all")
            ),
        )
        self._gainers_canvas_window = self._gainers_canvas.create_window(
            (0, 0), window=self._gainers_frame, anchor="nw"
        )
        self._gainers_canvas.configure(yscrollcommand=gainers_sb.set)

        def _on_gainers_canvas_resize(event):
            self._gainers_canvas.itemconfig(self._gainers_canvas_window,
                                            width=event.width)
        self._gainers_canvas.bind("<Configure>", _on_gainers_canvas_resize)

        self._gainers_canvas.pack(side="left", fill="both", expand=True)
        gainers_sb.pack(side="right", fill="y")

        right_panel = tk.Frame(main_body, bg=BG)
        right_panel.pack(side="left", fill="both", expand=True,
                         padx=(4, 4), pady=(6, 8))

        # Header card (draggable)
        header_card = tk.Frame(right_panel, bg=BG_CARD,
                               highlightbackground=BORDER, highlightthickness=1)
        header_card.pack(fill="x")
        header_card.bind("<Button-1>", self._start_drag)
        header_card.bind("<B1-Motion>", self._on_drag)

        header_inner = tk.Frame(header_card, bg=BG_CARD, padx=14, pady=4)
        header_inner.pack(fill="x")
        header_inner.bind("<Button-1>", self._start_drag)
        header_inner.bind("<B1-Motion>", self._on_drag)

        self.ticker_label = tk.Label(
            header_inner, text="Waiting...", fg=ACCENT,
            bg=BG_CARD, font=FONT_TICKER,
        )
        self.ticker_label.pack(side="left")

        self.company_label = tk.Label(
            header_inner, text="", fg=FG_DIM, bg=BG_CARD,
            font=FONT_UI, anchor="w",
        )
        self.company_label.pack(side="left", padx=(10, 0))

        self.overall_badge = tk.Label(
            header_inner, text="", fg="white", bg="#4A525C",
            font=FONT_UI_BOLD, padx=12, pady=6,
        )
        self.overall_badge.pack(side="right")

        self.history_badge = tk.Label(
            header_inner, text="", fg="white", bg="#4A525C",
            font=FONT_UI_BOLD, padx=12, pady=6,
        )
        self.history_badge.pack(side="right", padx=(0, 6))
        self.history_badge.pack_forget()  # hidden until data loaded

        # Badge: indica que los datos base vienen del caché de Top Gainers
        self._gainer_cache_badge = tk.Label(
            header_inner, text="⚡ TOP GAINER", fg="black", bg="#B9A816",
            font=FONT_MONO_BOLD, padx=8, pady=4,
        )
        self._gainer_cache_badge.pack(side="right", padx=(0, 6))
        self._gainer_cache_badge.pack_forget()  # hidden until data loaded

        # Line 1: MCap, Float/OS, Insiders, Inst.Own
        info_frame = tk.Frame(header_card, bg=BG_CARD)
        info_frame.pack(fill="x", padx=14, pady=(0, 2))
        self.info_label = tk.Label(
            info_frame, text="", fg=FG_INFO, bg=BG_CARD,
            font=FONT_UI, anchor="w",
        )
        self.info_label.pack(side="left")
        self._regsho_label = tk.Label(info_frame, text="", fg=FG_INFO, bg=BG_CARD, font=FONT_UI)
        self._regsho_label.pack(side="left")

        # Overhead warning label (blinking red)
        self._overhead_label = tk.Label(
            info_frame, text="⚠ TIENE OVERHEAD", fg="#FF3333", bg=BG_CARD,
            font=("Segoe UI Semibold", 10), padx=8,
        )
        self._overhead_label.pack(side="right")
        self._overhead_label.pack_forget()  # hidden until data loaded
        self._overhead_blink_job = None

        # Line 2: Sector | Country + flag
        info_frame2 = tk.Frame(header_card, bg=BG_CARD)
        info_frame2.pack(fill="x", padx=14, pady=(0, 10))
        self.info_label2 = tk.Label(
            info_frame2, text="", fg=FG_INFO, bg=BG_CARD,
            font=FONT_UI, anchor="w",
        )
        self.info_label2.pack(side="left")
        self._flag_label = tk.Label(info_frame2, bg=BG_CARD)
        self._flag_label.pack(side="left", padx=(2, 0))
        self._flag_photo = None  # keep reference to prevent GC

        # Scrollable content area
        container = tk.Frame(right_panel, bg=BG)
        container.pack(fill="both", expand=True, pady=(4, 0))

        canvas = tk.Canvas(container, bg=BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.content_frame = tk.Frame(canvas, bg=BG)

        self.content_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        self._canvas_window = canvas.create_window(
            (0, 0), window=self.content_frame, anchor="nw"
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_canvas_resize(event):
            canvas.itemconfig(self._canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_resize)

        # Mouse wheel scrolling — route to correct panel based on cursor position
        # Layout order (left to right): gainers | right panel
        def _on_mousewheel(event):
            x = event.x_root - self.root.winfo_rootx()
            if x < LEFT_PANEL_WIDTH + 12:
                self._gainers_canvas.yview_scroll(
                    int(-1 * (event.delta / 120)), "units"
                )
            else:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self._main_mousewheel_fn = _on_mousewheel
        self._main_canvas = canvas

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        # Re-bind when main window regains focus (secondary windows unbind_all on destroy)
        self.root.bind("<FocusIn>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.canvas = canvas

        self._show_waiting()

    # ── Display states ──────────────────────────────────────────────────────
    def _clear(self):
        for w in self.content_frame.winfo_children():
            w.destroy()

    def _show_waiting(self):
        self._clear()
        tk.Label(
            self.content_frame,
            text="Load a ticker in DAS or thinkorswim,\n"
                 "click a top gainer, or search above.",
            fg="#4A525C", bg=BG, font=("Segoe UI", 12), justify="center",
        ).pack(pady=60)

    def _update_history_badge(self, rating: str, post_url: str = ""):
        """Update the history badge in the header."""
        if rating in HISTORY_MAP:
            label, color = HISTORY_MAP[rating]
            self.history_badge.config(text=f"HISTORY: {label}", bg=color)
            self.history_badge.pack(side="right", padx=(0, 6))
            if post_url:
                self.history_badge.config(cursor="hand2")
                self.history_badge.bind("<Button-1>", lambda e, u=post_url: webbrowser.open(u))
            else:
                self.history_badge.config(cursor="")
                self.history_badge.unbind("<Button-1>")
        else:
            self.history_badge.pack_forget()

    def _update_overhead_label(self, has_overhead: bool):
        """Show or hide the blinking TIENE OVERHEAD label."""
        # Cancel any existing blink job
        if self._overhead_blink_job is not None:
            self.root.after_cancel(self._overhead_blink_job)
            self._overhead_blink_job = None

        if has_overhead:
            self._overhead_label.pack(side="right")
            self._overhead_label.config(fg="#FF3333")
            self._blink_overhead(visible=True)
        else:
            self._overhead_label.pack_forget()

    def _blink_overhead(self, visible: bool):
        """Toggle overhead label visibility for blink effect."""
        if not self._overhead_label.winfo_ismapped():
            return
        self._overhead_label.config(fg="#FF3333" if visible else BG_CARD)
        self._overhead_blink_job = self.root.after(500, self._blink_overhead, not visible)

    def _show_loading(self, ticker: str):
        self._clear()
        self.ticker_label.config(text=ticker)
        self.company_label.config(text="")
        self.overall_badge.config(text="...", bg="#4A525C")
        self.history_badge.pack_forget()
        self._gainer_cache_badge.pack_forget()
        self._update_overhead_label(False)
        self.info_label.config(text="Loading...")
        self.info_label2.config(text="")
        self._flag_label.config(image="", text="")
        self._flag_photo = None
        self._regsho_label.config(text="")
        tk.Label(
            self.content_frame,
            text=f"Fetching data for {ticker}...",
            fg=ACCENT, bg=BG, font=("Segoe UI", 12),
        ).pack(pady=60)
        self.root.update_idletasks()

    def _update_gainer_cache_badge(self, from_gainer_cache: bool):
        """Show or hide the TOP GAINER cache badge in the header."""
        if from_gainer_cache:
            self._gainer_cache_badge.pack(side="right", padx=(0, 6))
        else:
            self._gainer_cache_badge.pack_forget()

    def _show_no_data(self, ticker: str):
        self._clear()
        self.overall_badge.config(text="NO DATA", bg="#4A525C")
        self.info_label.config(text="")
        self.info_label2.config(text="")
        tk.Label(
            self.content_frame,
            text=f"No dilution data available for {ticker}.",
            fg="#FF6666", bg=BG, font=("Segoe UI", 11), justify="center",
        ).pack(pady=60)

    def _make_card(self, parent, title: str = None) -> tk.Frame:
        """Create a bordered card frame, optionally with a section header."""
        card = tk.Frame(parent, bg=BG_CARD,
                        highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x", padx=8, pady=(6, 0))
        if title:
            hdr = tk.Label(card, text=title, fg=ACCENT, bg=BG_CARD,
                           font=FONT_HEADER, anchor="w", padx=14, pady=10)
            hdr.pack(fill="x")
            tk.Frame(card, bg=BORDER, height=1).pack(fill="x")
        return card

    def _show_data(self, ticker: str, dilution: dict, floatdata: dict | None,
                   news: list[dict] | None = None, grok_line: str | None = None,
                   grok_date: str | None = None, grok_url: str | None = None,
                   in_play_warrants: list[dict] | None = None,
                   in_play_converts: list[dict] | None = None,
                   stock_price: float = 0.0,
                   jmt415_notes: list[dict] | None = None,
                   gap_stats: list[dict] | None = None,
                   offerings: list[dict] | None = None,
                   ownership: dict | None = None,
                   pump_dump: list[dict] | None = None,
                   research_tldr: list[dict] | None = None,
                   reverse_splits: list[dict] | None = None,
                   split_status: list[dict] | None = None,
                   open_compliance: list[dict] | None = None,
                   translated_commentary: str = "",
                   company_name: str = "",
                   flag_bytes: bytes | None = None,
                   claude_signal: dict | None = None):
        self._clear()

        # ── Store current data for on-demand ANALIZAR button ──
        self._current_data_for_claude = {
            "ticker": ticker,
            "company_name": company_name or "",
            "grok_text": grok_line or "",
            "news_headlines": [extract_headline(n) for n in (news or [])],
        }
        self._analizar_btn.config(text=" \u26a1 ANALIZAR ", bg="#FFD600", fg="black")

        # ── Add to recent tickers history ──
        from datetime import datetime as _dt
        self._ticker_history = [h for h in self._ticker_history if h["ticker"] != ticker]
        self._ticker_history.insert(0, {
            "ticker": ticker,
            "company_name": company_name or "",
            "timestamp": _dt.now().strftime("%H:%M"),
            "risk": (dilution or {}).get("overall_offering_risk", ""),
            "float": (floatdata or {}).get("tradable_float"),
        })
        if len(self._ticker_history) > 15:
            self._ticker_history = self._ticker_history[:15]

        dilution_url = f"https://app.askedgar.io/ticker/{ticker}/dilution"
        has_dilution = bool(dilution and dilution.get("overall_offering_risk"))

        # ── Claude LONG/SHORT signal card ──
        if claude_signal:
            signal    = claude_signal.get("signal", "")
            summary   = claude_signal.get("summary", "")
            catalysts = claude_signal.get("catalysts", [])
            action    = claude_signal.get("action", "")
            red_flags = claude_signal.get("red_flags", [])
            signal_color = "#2F7D57" if signal == "LONG" else "#A93232"
            dark_bg = "#0f1a14" if signal == "LONG" else "#1a0f0f"

            sig_card = tk.Frame(self.content_frame, bg=signal_color,
                                highlightbackground=signal_color, highlightthickness=2)
            sig_card.pack(fill="x", padx=8, pady=(6, 0))

            sig_inner = tk.Frame(sig_card, bg=dark_bg, padx=14, pady=12)
            sig_inner.pack(fill="x", padx=2, pady=(0, 2))

            # Header: $TICKER + company name + LONG/SHORT badge
            sig_top = tk.Frame(sig_inner, bg=dark_bg)
            sig_top.pack(fill="x")
            tk.Label(sig_top, text=f"${ticker}", fg="white", bg=dark_bg,
                     font=("Segoe UI Semibold", 14)).pack(side="left")
            if company_name:
                tk.Label(sig_top, text=f"  {company_name}", fg=FG_DIM, bg=dark_bg,
                         font=FONT_UI).pack(side="left")
            tk.Label(sig_top, text=f"  {signal}  ", fg="white", bg=signal_color,
                     font=("Segoe UI Semibold", 13, "bold"), padx=10, pady=4).pack(side="right")

            # Summary
            if summary:
                tk.Frame(sig_inner, bg=BORDER, height=1).pack(fill="x", pady=(10, 6))
                tk.Label(sig_inner, text="RESUMEN", fg=FG_DIM, bg=dark_bg,
                         font=("Consolas", 8)).pack(anchor="w")
                lbl_sum = tk.Label(sig_inner, text=summary, fg=FG, bg=dark_bg,
                                   font=FONT_UI, justify="left", anchor="w", wraplength=400)
                lbl_sum.pack(fill="x", pady=(4, 0))
                def _rw_sum(e, l=lbl_sum): l.config(wraplength=max(e.width-30,100))
                sig_inner.bind("<Configure>", _rw_sum)

            # Catalysts
            if catalysts:
                tk.Frame(sig_inner, bg=BORDER, height=1).pack(fill="x", pady=(10, 6))
                tk.Label(sig_inner, text="CATALIZADORES", fg=FG_DIM, bg=dark_bg,
                         font=("Consolas", 8)).pack(anchor="w")
                for cat in catalysts:
                    row = tk.Frame(sig_inner, bg=dark_bg)
                    row.pack(fill="x", pady=1)
                    tk.Label(row, text="◆", fg=ACCENT, bg=dark_bg,
                             font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))
                    lbl_cat = tk.Label(row, text=cat, fg=ACCENT, bg=dark_bg,
                                       font=FONT_UI, anchor="w", wraplength=380, justify="left")
                    lbl_cat.pack(side="left", fill="x", expand=True)

            # Action
            if action:
                tk.Frame(sig_inner, bg=BORDER, height=1).pack(fill="x", pady=(10, 6))
                tk.Label(sig_inner, text="ACCION RECOMENDADA", fg=FG_DIM, bg=dark_bg,
                         font=("Consolas", 8)).pack(anchor="w")
                action_box = tk.Frame(sig_inner, bg=signal_color,
                                      highlightbackground=signal_color, highlightthickness=1)
                action_box.pack(fill="x", pady=(6, 0))
                action_inner = tk.Frame(action_box, bg="#1a1a1a", padx=10, pady=8)
                action_inner.pack(fill="x", padx=1, pady=(0,1))
                # signal icon + label
                top_act = tk.Frame(action_inner, bg="#1a1a1a")
                top_act.pack(fill="x")
                tk.Label(top_act, text="📉" if signal=="SHORT" else "📈",
                         bg="#1a1a1a", font=("Segoe UI", 11)).pack(side="left")
                tk.Label(top_act, text=f"  {signal}", fg=signal_color, bg="#1a1a1a",
                         font=("Segoe UI Semibold", 11, "bold")).pack(side="left")
                lbl_act = tk.Label(action_inner, text=action, fg=FG, bg="#1a1a1a",
                                   font=FONT_UI, justify="left", anchor="w", wraplength=380)
                lbl_act.pack(fill="x", pady=(6, 0))
                def _rw_act(e, l=lbl_act): l.config(wraplength=max(e.width-20,100))
                action_inner.bind("<Configure>", _rw_act)

            # Red flags
            if red_flags:
                tk.Frame(sig_inner, bg=BORDER, height=1).pack(fill="x", pady=(10, 6))
                tk.Label(sig_inner, text="RED FLAGS", fg=FG_DIM, bg=dark_bg,
                         font=("Consolas", 8)).pack(anchor="w")
                for flag in red_flags:
                    row = tk.Frame(sig_inner, bg=dark_bg)
                    row.pack(fill="x", pady=1)
                    tk.Label(row, text="🚩", bg=dark_bg,
                             font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))
                    lbl_flag = tk.Label(row, text=flag, fg=RED, bg=dark_bg,
                                        font=FONT_UI, anchor="w", wraplength=380, justify="left")
                    lbl_flag.pack(side="left", fill="x", expand=True)

        if has_dilution:
            risk = dilution.get("overall_offering_risk", "N/A")
            self.overall_badge.config(text=f"RISK: {risk}", bg=risk_bg(risk))
        else:
            self.overall_badge.config(text="NO DATA", bg="#4A525C")

        # ── Company name ──
        self.company_label.config(text=company_name if company_name else "")

        # ── Info lines from screener + ownership data ──
        if floatdata:
            flt  = fmt_millions(floatdata.get("tradable_float"))
            outs = fmt_millions(floatdata.get("outstanding"))
            mc   = fmt_millions(floatdata.get("market_cap"))
            sector  = floatdata.get("sector", "")
            country = floatdata.get("country", "")

            # Insiders & institutional from ownership
            insider_pct = ""
            inst_pct    = ""
            if ownership and ownership.get("owners"):
                insider_shares = 0
                inst_shares    = 0
                for o in ownership["owners"]:
                    otype  = (o.get("owner_type") or o.get("title") or "").lower()
                    shares = o.get("common_shares_amount") or 0
                    if "insider" in otype or "director" in otype or "officer" in otype:
                        insider_shares += shares
                    elif "institution" in otype or "fund" in otype or "invest" in otype:
                        inst_shares += shares
                outstanding_raw = floatdata.get("outstanding")
                if outstanding_raw and outstanding_raw > 0:
                    if insider_shares > 0:
                        insider_pct = f"{insider_shares / outstanding_raw * 100:.1f}%"
                    if inst_shares > 0:
                        inst_pct = f"{inst_shares / outstanding_raw * 100:.1f}%"

            # Line 1: MCap | Float/OS | Insiders | Inst.Own
            parts1 = [f"Mkt Cap: {mc}", f"Float/OS: {flt}/{outs}"]
            if insider_pct:
                parts1.append(f"Insiders: {insider_pct}")
            if inst_pct:
                parts1.append(f"Inst. Own: {inst_pct}")
            if dilution.get("regsho"):
                parts1.append("Reg SHO")
            self.info_label.config(text="  |  ".join(parts1))
            self._regsho_label.config(text="")

            # Line 2: Sector | Country + flag
            parts2 = []
            if sector:
                parts2.append(sector)
            if country:
                parts2.append(country + " ")
            self.info_label2.config(text="  |  ".join(parts2))
            if flag_bytes:
                try:
                    from PIL import Image, ImageTk
                    import io
                    img = Image.open(io.BytesIO(flag_bytes))
                    self._flag_photo = ImageTk.PhotoImage(img)
                    self._flag_label.config(image=self._flag_photo, text="")
                except Exception:
                    self._flag_label.config(image="", text="")
                    self._flag_photo = None
            else:
                self._flag_label.config(image="", text="")
                self._flag_photo = None
        else:
            self.info_label.config(text="")
            self.info_label2.config(text="")
            self._flag_label.config(image="", text="")
            self._flag_photo = None
            self._regsho_label.config(text="")

        # ── Feed card (news + grok) ──
        has_feed = news or grok_line
        if has_feed:
            feed_card = self._make_card(self.content_frame)
            feed_inner = tk.Frame(feed_card, bg=BG_CARD, padx=8, pady=8)
            feed_inner.pack(fill="x")

            if news:
                for item in news:
                    headline = item.get("_translated_headline") or extract_headline(item)
                    url = (item.get("url") or item.get("document_url") or
                           f"https://app.askedgar.io/ticker/{ticker}/news")
                    form = item.get("form_type", "")
                    raw_date = item.get("created_at") or item.get("filed_at", "")
                    et_date, spain_date = convert_api_date(raw_date)
                    self._add_feed_item(feed_inner, form, headline, url, et_date, spain_date)

            if grok_line:
                et_date, spain_date = "", ""
                if grok_date:
                    et_date, spain_date = convert_api_date(grok_date)
                self._add_feed_item(feed_inner, "grok", grok_line, grok_url, et_date, spain_date)

        # ── Risk badges card (grid: 3 columns, wraps to 2 rows) ──
        if has_dilution:
            risk = dilution.get("overall_offering_risk", "N/A")
            badges_card = self._make_card(self.content_frame)
            badges_inner = tk.Frame(badges_card, bg=BG_CARD, padx=8, pady=8, cursor="hand2")
            badges_inner.pack(fill="x")
            badges_inner.bind("<Button-1>", lambda e, u=dilution_url: webbrowser.open(u))

            badge_items = [
                ("Overall Risk", risk),
                ("Offering", dilution.get("offering_ability", "N/A")),
                ("Dilution", dilution.get("dilution", "N/A")),
                ("Frequency", dilution.get("offering_frequency", "N/A")),
                ("Cash Need", dilution.get("cash_need", "N/A")),
                ("Warrants", dilution.get("warrant_exercise", "N/A")),
            ]
            for i, (label, level) in enumerate(badge_items):
                self._add_badge_grid(badges_inner, label, level, dilution_url,
                                     row=i // 3, col=i % 3)
            badges_inner.columnconfigure((0, 1, 2), weight=1)

        # ── Offering Ability card ──
        offering_desc = dilution.get("offering_ability_desc") if has_dilution else None
        if offering_desc:
            self._add_offering_ability_card(offering_desc, url=dilution_url)

        # ── In Play Dilution card ──
        if in_play_warrants or in_play_converts:
            self._add_in_play_section(in_play_warrants or [], in_play_converts or [], stock_price, dilution_url)

        # ── NASDAQ Compliance card (only OPEN deficiencies, pre-filtered) ──
        if open_compliance:
            self._add_nasdaq_compliance_card(open_compliance)

        # ── Recent Offerings card ──
        if offerings:
            self._add_offerings_card(offerings[:3], stock_price, url=dilution_url)

        # ── Últimas 10 Noticias card ──
        all_news_raw = _cached_news_results(ticker)
        last10 = [r for r in (all_news_raw or []) if r.get("form_type") in ("news", "8-K", "6-K")][:10]
        if last10:
            self._add_last_news_card(last10, ticker)

        # ── Gap Stats card ──
        if gap_stats:
            self._add_gap_stats_card(gap_stats)

        # ── Splits card (reverse splits + split status) ──
        if reverse_splits or split_status:
            self._add_splits_card(reverse_splits or [], split_status or [])

        # ── JMT415 Previous Notes card ──
        if jmt415_notes:
            self._add_jmt415_card(jmt415_notes)

        # ── Management Commentary card ──
        if translated_commentary:
            self._add_section_card("Mgmt Commentary", translated_commentary, url=dilution_url)

        # ── Ownership card ──
        if ownership and ownership.get("owners"):
            self._add_ownership_card(ownership)

        # ── Pump & Dump Tracker card ──
        if pump_dump:
            self._add_pump_dump_card(pump_dump)

        # ── Research Report TLDR card (translated only) ──
        if research_tldr:
            self._add_research_tldr_card(research_tldr)

        # (NASDAQ Compliance is shown above, between In Play Dilution and Recent Offerings)

        # ── Mis Notas card ──
        self._add_notes_card(ticker)



    def _add_notes_card(self, ticker: str):
        """Notas personales por ticker — editable inline, guardado automático."""
        existing_text = read_ticker_notes(ticker)

        card = tk.Frame(self.content_frame, bg=BG_CARD,
                        highlightbackground=BORDER_ACCENT, highlightthickness=1)
        card.pack(fill="x", padx=8, pady=(6, 0))

        # Header row
        hdr = tk.Frame(card, bg=BG_CARD, padx=14, pady=8)
        hdr.pack(fill="x")

        tk.Label(hdr, text="📝  MIS NOTAS", fg=ACCENT, bg=BG_CARD,
                 font=FONT_HEADER).pack(side="left")

        self._notes_status_lbl = tk.Label(hdr, text="", fg=FG_DIM, bg=BG_CARD,
                                          font=FONT_MONO)
        self._notes_status_lbl.pack(side="right")

        save_btn = tk.Label(hdr, text="  GUARDAR  ", fg=BG, bg=ACCENT,
                            font=FONT_UI_BOLD, padx=8, pady=2, cursor="hand2")
        save_btn.pack(side="right", padx=(0, 6))

        clear_btn = tk.Label(hdr, text="  BORRAR  ", fg="white", bg="#A93232",
                             font=FONT_UI_BOLD, padx=8, pady=2, cursor="hand2")
        clear_btn.pack(side="right", padx=(0, 6))

        tk.Frame(card, bg=BORDER, height=1).pack(fill="x")

        # Text area
        text_frame = tk.Frame(card, bg=BG_CARD, padx=10, pady=8)
        text_frame.pack(fill="x")

        txt = tk.Text(
            text_frame,
            bg=BG_ROW, fg=FG, insertbackground=FG,
            font=FONT_UI, relief="flat",
            height=5, wrap="word",
            highlightbackground=BORDER_INNER, highlightthickness=1,
            padx=8, pady=6,
        )
        txt.pack(fill="x")
        if existing_text:
            txt.insert("1.0", existing_text)

        def _save(event=None):
            notes = txt.get("1.0", "end-1c").strip()
            write_ticker_notes(ticker, notes)
            self._notes_status_lbl.config(text="✓ guardado", fg="#4CAF50")
            self.root.after(2000, lambda: self._notes_status_lbl.config(text="", fg=FG_DIM))

        def _clear():
            txt.delete("1.0", "end")
            write_ticker_notes(ticker, "")
            self._notes_status_lbl.config(text="✓ borrado", fg=RED)
            self.root.after(2000, lambda: self._notes_status_lbl.config(text="", fg=FG_DIM))

        save_btn.bind("<Button-1>", _save)
        clear_btn.bind("<Button-1>", lambda e: _clear())
        # Ctrl+S also saves
        txt.bind("<Control-s>", _save)

        # Auto-save on focus out
        txt.bind("<FocusOut>", _save)

        # Hint text if empty
        if not existing_text:
            HINT = "Escribe tus notas aquí... (se guardan automáticamente)"
            txt.insert("1.0", HINT)
            txt.config(fg="#4A525C")

            def _on_focus_in(e):
                if txt.get("1.0", "end-1c") == HINT:
                    txt.delete("1.0", "end")
                    txt.config(fg=FG)

            def _on_focus_out_hint(e):
                if not txt.get("1.0", "end-1c").strip():
                    txt.insert("1.0", HINT)
                    txt.config(fg="#4A525C")
                else:
                    _save()

            txt.bind("<FocusIn>", _on_focus_in)
            txt.bind("<FocusOut>", _on_focus_out_hint)

    def _add_badge_grid(self, parent, label: str, level: str,
                        url: str | None = None, row: int = 0, col: int = 0):
        """Place a badge in a grid layout (3 columns, rows wrap automatically)."""
        frame = tk.Frame(parent, bg=BG_CARD, padx=4, pady=4, cursor="hand2")
        frame.grid(row=row, column=col, padx=4, pady=2, sticky="ew")

        lbl = tk.Label(
            frame, text=label, fg=FG_DIM, bg=BG_CARD,
            font=FONT_MONO, cursor="hand2",
        )
        lbl.pack()

        badge = tk.Label(
            frame, text=f" {level} ", fg="white", bg=risk_bg(level),
            font=FONT_UI_BOLD, padx=8, pady=3, cursor="hand2",
        )
        badge.pack()

        if url:
            for w in (frame, lbl, badge):
                w.bind("<Button-1>", lambda e, u=url: webbrowser.open(u))

    def _add_feed_item(self, parent, form_type: str, headline: str,
                       url: str | None, date: str = "", spain_date: str = ""):
        """Feed row with source stripe on the left. Entire row is clickable."""
        SOURCE_COLORS = {
            "news": "#1F8FB3",
            "8-K": "#A85C14",
            "6-K": "#A85C14",
            "grok": "#7B3FA0",
        }
        source_color = SOURCE_COLORS.get(form_type, "#555555")
        tag = form_type.upper() if form_type != "news" else "NEWS"

        # Truncate grok output to ~240 chars
        if form_type == "grok" and len(headline) > 240:
            headline = headline[:237] + "..."

        row = tk.Frame(parent, bg=BG_ROW,
                       highlightbackground=BORDER_INNER, highlightthickness=1)
        row.pack(fill="x", pady=2)

        # Source stripe (left column)
        stripe = tk.Label(
            row, text=tag, fg="white", bg=source_color,
            font=("Consolas", 8, "bold"), width=5, padx=4, pady=8,
        )
        stripe.pack(side="left", fill="y")

        # Content area — stacked vertically so text wraps downward
        content = tk.Frame(row, bg=BG_ROW, padx=10, pady=6)
        content.pack(side="left", fill="both", expand=True)

        if date:
            date_frame = tk.Frame(content, bg=BG_ROW)
            date_frame.pack(fill="x")
            tk.Label(date_frame, text=date + ",", fg=FG_DIM, bg=BG_ROW,
                     font=FONT_MONO, anchor="w").pack(side="left")
            flag = get_spain_flag_photo()
            if flag:
                tk.Label(date_frame, image=flag, bg=BG_ROW).pack(side="left", padx=(4, 2))
            if spain_date:
                tk.Label(date_frame, text=f"({spain_date})", fg=FG_DIM, bg=BG_ROW,
                         font=FONT_MONO).pack(side="left")

        hl_label = tk.Label(
            content, text=headline, fg="white", bg=BG_ROW,
            font=FONT_UI_BOLD, anchor="w", wraplength=200,
            justify="left",
        )
        hl_label.pack(fill="x")

        def _rewrap_hl(event, lbl=hl_label):
            lbl.config(wraplength=max(event.width - 30, 100))
        content.bind("<Configure>", _rewrap_hl)

        # Make entire row clickable if there's a URL
        if url:
            row.config(cursor="hand2")
            def _bind_click(widget, target_url):
                widget.bind("<Button-1>", lambda e, u=target_url: webbrowser.open(u))
                widget.config(cursor="hand2")
            for w in (row, stripe, content, hl_label):
                _bind_click(w, url)
            for child in content.winfo_children():
                _bind_click(child, url)

    def _bind_card_click(self, card, url: str):
        """Make an entire card and all its descendants clickable."""
        def _bind(w, u=url):
            w.bind("<Button-1>", lambda e, u=u: webbrowser.open(u))
            w.config(cursor="hand2")
        def _bind_all(widget):
            _bind(widget)
            for child in widget.winfo_children():
                _bind_all(child)
        _bind_all(card)

    def _add_section_card(self, title: str, text: str, url: str = ""):
        """Section card with header + bottom border + wrapped text content."""
        card = self._make_card(self.content_frame, title=title)
        body = tk.Frame(card, bg=BG_CARD, padx=14, pady=14)
        body.pack(fill="x")
        text_label = tk.Label(
            body, text=text, fg=FG, bg=BG_CARD,
            font=FONT_UI, justify="left", anchor="w",
        )
        text_label.pack(fill="x")
        def _rewrap(event, lbl=text_label):
            lbl.config(wraplength=max(event.width - 4, 100))
        body.bind("<Configure>", _rewrap)
        if url:
            self._bind_card_click(card, url)

    def _add_offering_ability_card(self, desc: str, url: str = ""):
        """Offering Ability card with color-coded capacity values."""
        card = self._make_card(self.content_frame, title="Offering Ability")
        body = tk.Frame(card, bg=BG_CARD, padx=14, pady=14)
        body.pack(fill="x")

        # Parse and color individual segments — stacked vertically
        parts = [p.strip() for p in desc.split(",")]

        for part in parts:
            part_lower = part.lower()
            if "pending s-1" in part_lower or "pending f-1" in part_lower:
                color = "#4CAF50"
                bold = True
            elif ("shelf capacity" in part_lower or "atm capacity" in part_lower
                  or "equity line capacity" in part_lower):
                if "$0.00" in part:
                    color = "#FF4444"
                    bold = False
                else:
                    color = "#4CAF50"
                    bold = True
            else:
                color = FG
                bold = False

            font = ("Segoe UI Semibold", 10) if bold else FONT_UI
            tk.Label(
                body, text=part, fg=color, bg=BG_CARD,
                font=font, anchor="w",
            ).pack(fill="x")

        if url:
            self._bind_card_click(card, url)

    def _add_last_news_card(self, news_items: list[dict], ticker: str):
        """Card showing the last 10 news/8-K/6-K items for the ticker."""
        card = self._make_card(self.content_frame, title="Últimas Noticias (10)")
        body = tk.Frame(card, bg=BG_CARD, padx=10, pady=8)
        body.pack(fill="x")

        TYPE_COLORS = {
            "news": "#1F8FB3",
            "8-K":  "#B96A16",
            "6-K":  "#B96A16",
        }

        for i, item in enumerate(news_items):
            form = item.get("form_type", "news")
            headline = item.get("_translated_headline") or extract_headline(item)
            raw_date = item.get("created_at") or item.get("filed_at", "")
            date_display, time_display = convert_api_date(raw_date)
            url = (item.get("url") or item.get("document_url") or
                   f"https://app.askedgar.io/ticker/{ticker}/news")
            stripe_color = TYPE_COLORS.get(form, "#555555")

            row_bg = BG_ROW if i % 2 == 0 else BG_ROW_ALT
            row = tk.Frame(body, bg=row_bg,
                           highlightbackground=BORDER_INNER, highlightthickness=1)
            row.pack(fill="x", pady=2)

            # Left color stripe
            stripe = tk.Frame(row, bg=stripe_color, width=3)
            stripe.pack(side="left", fill="y")

            inner = tk.Frame(row, bg=row_bg, padx=10, pady=6)
            inner.pack(side="left", fill="x", expand=True)

            # Top line: form type badge + date + time
            top = tk.Frame(inner, bg=row_bg)
            top.pack(fill="x")
            tk.Label(top, text=f" {form} ", fg="white", bg=stripe_color,
                     font=("Consolas", 7, "bold"), padx=4, pady=1).pack(side="left")
            if date_display:
                tk.Label(top, text=f"  {date_display}", fg=FG_DIM, bg=row_bg,
                         font=FONT_MONO).pack(side="left")
            if time_display:
                tk.Label(top, text=f"  {time_display}", fg=FG_DIM, bg=row_bg,
                         font=FONT_MONO).pack(side="left")

            # Headline
            lbl = tk.Label(inner, text=headline, fg=FG, bg=row_bg,
                           font=FONT_UI, anchor="w", justify="left",
                           wraplength=350, cursor="hand2")
            lbl.pack(fill="x", pady=(3, 0))

            def _rewrap(event, l=lbl):
                l.config(wraplength=max(event.width - 20, 100))
            inner.bind("<Configure>", _rewrap)

            # Click to open URL
            for w in (row, stripe, inner, top, lbl):
                w.bind("<Button-1>", lambda e, u=url: webbrowser.open(u))

    def _add_gap_stats_card(self, gaps: list[dict]):
        """Gap Stats summary card."""
        from datetime import datetime
        card = self._make_card(self.content_frame, title="Gap Stats")
        body = tk.Frame(card, bg=BG_CARD, padx=14, pady=10)
        body.pack(fill="x")

        n = len(gaps)
        last_date = gaps[0].get("date", "N/A") if gaps else "N/A"

        # Compute averages
        gap_pcts = [g["gap_percentage"] for g in gaps if g.get("gap_percentage") is not None]
        avg_gap = sum(gap_pcts) / len(gap_pcts) if gap_pcts else 0

        oh_spikes = []
        ol_drops = []
        for g in gaps:
            o = g.get("market_open")
            h = g.get("high_price")
            lo = g.get("low_price")
            if o and o > 0:
                if h is not None:
                    oh_spikes.append((h - o) / o * 100)
                if lo is not None:
                    ol_drops.append((lo - o) / o * 100)

        avg_oh = sum(oh_spikes) / len(oh_spikes) if oh_spikes else 0
        avg_ol = sum(ol_drops) / len(ol_drops) if ol_drops else 0

        # % new high after 11am EST (high_time is already EST, e.g. "2026-03-27T12:34:00")
        high_after_11 = 0
        for g in gaps:
            ht = g.get("high_time", "")
            if ht:
                try:
                    t = datetime.fromisoformat(ht)
                    if t.hour >= 11:
                        high_after_11 += 1
                except Exception:
                    pass
        pct_high_after_11 = (high_after_11 / n * 100) if n else 0

        # % closed below VWAP (API gives closed_over_vwap boolean)
        below_vwap = sum(1 for g in gaps if g.get("closed_over_vwap") is False)
        pct_below_vwap = (below_vwap / n * 100) if n else 0

        # % closed below open
        below_open = sum(1 for g in gaps if g.get("market_close") and g.get("market_open")
                         and g["market_close"] < g["market_open"])
        pct_below_open = (below_open / n * 100) if n else 0

        # Avg premarket volume and avg volume
        pm_vols = [g["premarket_volume"] for g in gaps if g.get("premarket_volume")]
        avg_pm_vol = sum(pm_vols) / len(pm_vols) if pm_vols else 0
        vols = [g["volume"] for g in gaps if g.get("volume")]
        avg_vol = sum(vols) / len(vols) if vols else 0

        # Display stats as label-value rows
        stats = [
            ("Last Gap Date", last_date),
            ("Number of Gaps", str(n)),
            ("Avg Gap %", f"{avg_gap:.1f}%"),
            ("Avg Open→High", f"+{avg_oh:.1f}%"),
            ("Avg Open→Low", f"{avg_ol:.1f}%"),
            ("New High After 11am", f"{pct_high_after_11:.0f}%"),
            ("Closed Below VWAP", f"{pct_below_vwap:.0f}%"),
            ("Closed Below Open", f"{pct_below_open:.0f}%"),
            ("Avg PreMarket Vol", fmt_volume(avg_pm_vol)),
            ("Avg Volume", fmt_volume(avg_vol)),
        ]

        for label, value in stats:
            row = tk.Frame(body, bg=BG_CARD)
            row.pack(fill="x", pady=1)
            tk.Label(row, text=label, fg=FG_DIM, bg=BG_CARD,
                     font=FONT_MONO, width=22, anchor="w").pack(side="left")
            # Color code certain values
            ORANGE = "#B96A16"
            val_color = FG
            if "Below VWAP" in label:
                try:
                    pv = float(value.rstrip("%"))
                    val_color = GREEN if pv <= 59 else (ORANGE if pv <= 84 else RED)
                except ValueError:
                    pass
            elif "Below Open" in label:
                try:
                    pv = float(value.rstrip("%"))
                    val_color = GREEN if pv <= 50 else (ORANGE if pv <= 74 else RED)
                except ValueError:
                    pass
            elif "After 11am" in label:
                try:
                    pv = float(value.rstrip("%"))
                    val_color = GREEN if pv >= 45 else (ORANGE if pv >= 21 else RED)
                except ValueError:
                    pass
            elif "Open→High" in label:
                val_color = GREEN
            elif "Open→Low" in label:
                val_color = RED
            tk.Label(row, text=value, fg=val_color, bg=BG_CARD,
                     font=FONT_MONO_BOLD, anchor="w").pack(side="left")


    def _add_offerings_card(self, offerings: list[dict], stock_price: float = 0.0,
                            url: str = ""):
        """Recent Offerings card with headline + data row per offering."""
        card = self._make_card(self.content_frame, title="Recent Offerings")
        body = tk.Frame(card, bg=BG_CARD, padx=14, pady=10)
        body.pack(fill="x")

        for i, o in enumerate(offerings):
            row_bg = BG_ROW if i % 2 == 0 else BG_ROW_ALT

            row = tk.Frame(body, bg=row_bg,
                           highlightbackground=BORDER_INNER, highlightthickness=1)
            row.pack(fill="x", pady=2)

            inner = tk.Frame(row, bg=row_bg, padx=10, pady=6)
            inner.pack(fill="x")

            headline = (o.get("headline") or "Offering").strip()
            tk.Label(inner, text=headline, fg="white", bg=row_bg,
                     font=FONT_UI, anchor="w").pack(fill="x")

            is_atm = "ATM USED" in headline.upper()

            data_row = tk.Frame(inner, bg=row_bg)
            data_row.pack(fill="x", pady=(2, 0))

            if is_atm:
                offering_amt = o.get("offering_amount")
                if offering_amt:
                    tk.Label(data_row, text=f"${fmt_millions(offering_amt)}", fg="#4CAF50", bg=row_bg,
                             font=FONT_MONO_BOLD).pack(side="left")
                filed = (o.get("filed_at") or "")[:10]
                if filed:
                    tk.Label(data_row, text="  |  ", fg=FG_DIM, bg=row_bg,
                             font=FONT_MONO).pack(side="left")
                    tk.Label(data_row, text=filed, fg=FG_DIM, bg=row_bg,
                             font=FONT_MONO).pack(side="left")
            else:
                offer_price = o.get("share_price") or 0
                in_money = stock_price > 0 and offer_price > 0 and offer_price <= stock_price
                highlight = "#4CAF50" if in_money else "#FF9800"

                shares = o.get("shares_amount")
                warrants = o.get("warrants_amount")
                filed = (o.get("filed_at") or "")[:10]

                parts_colored = []
                if shares:
                    parts_colored.append(f"Amt:{fmt_millions(shares)}")
                if offer_price:
                    parts_colored.append(f"${offer_price:.2f}")
                if warrants:
                    parts_colored.append(f"Wrrnts:{fmt_millions(warrants)}")

                for j, part in enumerate(parts_colored):
                    if j > 0:
                        tk.Label(data_row, text=" | ", fg=FG_DIM, bg=row_bg,
                                 font=FONT_MONO).pack(side="left")
                    tk.Label(data_row, text=part, fg=highlight, bg=row_bg,
                             font=FONT_MONO_BOLD).pack(side="left")

                if filed:
                    if parts_colored:
                        tk.Label(data_row, text="  |  ", fg=FG_DIM, bg=row_bg,
                                 font=FONT_MONO).pack(side="left")
                    tk.Label(data_row, text=filed, fg=FG_DIM, bg=row_bg,
                             font=FONT_MONO).pack(side="left")

        if url:
            self._bind_card_click(card, url)

    def _add_jmt415_card(self, notes: list[dict]):
        """JMT415 Previous Notes card with bordered panels per note."""
        card = self._make_card(self.content_frame, title="JMT415 Previous Notes")
        body = tk.Frame(card, bg=BG_CARD, padx=10, pady=10)
        body.pack(fill="x")

        for i, note in enumerate(notes):
            date = (note.get("filed_at") or "")[:10]
            text = note.get("_translated_text") or (note.get("summary") or note.get("title") or "Note").strip()
            row_bg = BG_ROW if i % 2 == 0 else BG_ROW_ALT

            row = tk.Frame(body, bg=row_bg,
                           highlightbackground=BORDER_INNER, highlightthickness=1)
            row.pack(fill="x", pady=2)

            inner = tk.Frame(row, bg=row_bg, padx=10, pady=8)
            inner.pack(fill="x")

            tk.Label(inner, text=date, fg=FG_DIM, bg=row_bg,
                     font=FONT_MONO).pack(anchor="w")
            note_label = tk.Label(inner, text=text, fg=FG, bg=row_bg,
                                  font=FONT_UI, anchor="w",
                                  wraplength=350, justify="left")
            note_label.pack(fill="x", pady=(2, 0))

            def _rewrap(event, lbl=note_label):
                lbl.config(wraplength=max(event.width - 40, 100))
            row.bind("<Configure>", _rewrap)

    def _add_ownership_card(self, ownership: dict):
        """Ownership card showing latest reported date with owner table."""
        reported_date = (ownership.get("reported_date") or "")[:10]
        title = f"Ownership  ({reported_date})" if reported_date else "Ownership"
        card = self._make_card(self.content_frame, title=title)
        body = tk.Frame(card, bg=BG_CARD, padx=10, pady=10)
        body.pack(fill="x")

        # Table header
        hdr = tk.Frame(body, bg=BG_CARD)
        hdr.pack(fill="x", pady=(0, 4))
        tk.Label(hdr, text="Owner", fg=ACCENT, bg=BG_CARD,
                 font=FONT_MONO_BOLD, anchor="w", width=20).pack(side="left")
        tk.Label(hdr, text="Title", fg=ACCENT, bg=BG_CARD,
                 font=FONT_MONO_BOLD, anchor="w", width=14).pack(side="left")
        tk.Label(hdr, text="Shares", fg=ACCENT, bg=BG_CARD,
                 font=FONT_MONO_BOLD, anchor="e").pack(side="right")

        doc_url = ""
        owners = ownership.get("owners", [])
        for i, owner in enumerate(owners):
            row_bg = BG_ROW if i % 2 == 0 else BG_ROW_ALT
            row = tk.Frame(body, bg=row_bg)
            row.pack(fill="x", pady=1)
            inner = tk.Frame(row, bg=row_bg, padx=6, pady=4)
            inner.pack(fill="x")

            name = owner.get("owner_name", "")
            title_str = owner.get("title", "") or owner.get("owner_type", "")
            shares = owner.get("common_shares_amount", 0)
            shares_str = f"{shares:,.0f}" if shares else "0"

            tk.Label(inner, text=name, fg=FG, bg=row_bg,
                     font=FONT_MONO, anchor="w", width=20).pack(side="left")
            tk.Label(inner, text=title_str, fg=FG_DIM, bg=row_bg,
                     font=FONT_MONO, anchor="w", width=14).pack(side="left")
            tk.Label(inner, text=shares_str, fg="#4CAF50", bg=row_bg,
                     font=FONT_MONO_BOLD, anchor="e").pack(side="right")

            if not doc_url:
                doc_url = owner.get("document_url", "")

        if doc_url:
            self._bind_card_click(card, doc_url)

    def _add_pump_dump_card(self, results: list[dict]):
        """Pump & Dump Tracker card with risk badges."""
        card = self._make_card(self.content_frame, title="Pump & Dump Tracker")
        body = tk.Frame(card, bg=BG_CARD, padx=10, pady=10)
        body.pack(fill="x")

        item = results[0]

        # Risk badges row
        risk_fields = [
            ("Country Risk", "country_risk"),
            ("Scam Risk", "scam_risk"),
            ("Float Risk", "float_risk"),
            ("Underwriter Risk", "underwriter_risk"),
        ]
        badge_frame = tk.Frame(body, bg=BG_CARD)
        badge_frame.pack(fill="x", pady=(0, 8))
        for col_idx, (label, key) in enumerate(risk_fields):
            level = (item.get(key) or "N/A").capitalize()
            cell = tk.Frame(badge_frame, bg=BG_CARD)
            cell.grid(row=0, column=col_idx, sticky="ew", padx=2)
            tk.Label(cell, text=label, fg=FG_DIM, bg=BG_CARD,
                     font=("Consolas", 8)).pack()
            tk.Label(cell, text=level, fg="white", bg=risk_bg(level),
                     font=("Consolas", 9, "bold"), padx=8, pady=2).pack()
        badge_frame.columnconfigure((0, 1, 2, 3), weight=1)

        # Key data fields
        data_fields = [
            ("IPO Date", (item.get("ipo_date") or "")[:10]),
            ("Lock-Up Exp.", (item.get("lock_up_expiration") or "")[:10]),
            ("Underwriters", item.get("underwriters") or "N/A"),
            ("Liquidations", str(item.get("number_liquidations") or "N/A")),
            ("Last Liquidation", (item.get("last_liquidation_date") or "")[:10]),
            ("Country", item.get("country") or "N/A"),
            ("ADR", "Yes" if item.get("isadr") else "No"),
            ("Gain 1D", f"{item.get('gain_1_day', 0) or 0:.1f}%"),
            ("Gain 7D", f"{item.get('gain_7_day', 0) or 0:.1f}%"),
            ("Gain 30D", f"{item.get('gain_30_day', 0) or 0:.1f}%"),
        ]
        for i, (label, value) in enumerate(data_fields):
            if not value or value == "N/A" and label in ("IPO Date", "Lock-Up Exp.", "Last Liquidation"):
                continue
            row_bg = BG_ROW if i % 2 == 0 else BG_ROW_ALT
            row = tk.Frame(body, bg=row_bg)
            row.pack(fill="x", pady=1)
            inner = tk.Frame(row, bg=row_bg, padx=8, pady=3)
            inner.pack(fill="x")
            tk.Label(inner, text=label, fg=FG_DIM, bg=row_bg,
                     font=FONT_MONO, anchor="w", width=18).pack(side="left")
            val_color = FG
            if "Gain" in label:
                try:
                    val_color = GREEN if float(value.replace("%", "")) >= 0 else RED
                except ValueError:
                    pass
            tk.Label(inner, text=value, fg=val_color, bg=row_bg,
                     font=FONT_MONO_BOLD, anchor="e").pack(side="right")

    def _parse_tldr_line(self, line: str) -> tuple[str, str, str, list[str]]:
        """Parse a TLDR line into (signal_color, date_prefix, text_to_translate, [urls])."""
        import re as _re
        line = line.strip()
        # Remove leading bullet
        if line.startswith("•"):
            line = line[1:].strip()

        # Detect signal emoji and assign color
        SIGNAL_COLORS = {
            "\U0001f7e2": "#4CAF50",   # 🟢 green
            "\U0001f7e1": "#FFD600",   # 🟡 yellow
            "\U0001f534": "#FF4444",   # 🔴 red
        }
        signal_color = ""
        for emoji, color in SIGNAL_COLORS.items():
            if emoji in line:
                signal_color = color
                line = line.replace(emoji, "").strip()
                break

        # Extract date prefix (e.g. "09-Apr-26:", "H2-26:", "30-Apr-2026:")
        date_prefix = ""
        date_match = _re.match(r'^(\d{1,2}-\w{3}-\d{2,4}\s*:|H\d-\d{2,4}\s*:)', line)
        if date_match:
            date_prefix = date_match.group(1)
            line = line[date_match.end():].strip()

        # Extract URLs
        urls = _re.findall(r'<(https?://[^>]+)>', line)
        if not urls:
            urls = _re.findall(r'(https?://\S+)', line)
        # Remove URLs from text
        clean = _re.sub(r'<https?://[^>]+>', '', line)
        clean = _re.sub(r'https?://\S+', '', clean)
        clean = clean.strip()

        return signal_color, date_prefix, clean, urls

    def _add_research_tldr_card(self, results: list[dict]):
        """Research TLDR card — formatted with colored signals, translated text, clickable links."""
        if not results:
            return
        item = results[0]
        text = (item.get("tldr_text") or item.get("analysis_text") or
                item.get("summary") or item.get("report") or item.get("text") or "")
        if not text:
            for v in item.values():
                if isinstance(v, str) and len(v) > 50:
                    text = v
                    break
        if not text:
            return

        date = (item.get("created_at") or item.get("last_updated") or "")[:10]
        post_url = item.get("post_url", "")

        header = "Research Report (TLDR)"
        if date:
            header += f"  ({date})"

        card = self._make_card(self.content_frame, title=header)
        body = tk.Frame(card, bg=BG_CARD, padx=14, pady=10)
        body.pack(fill="x")

        # Use pre-translated lines if available, otherwise parse and show raw
        pre_translated = item.get("_translated_lines")
        if pre_translated:
            parsed_lines = pre_translated
        else:
            parsed_lines = []
            for ln in text.split("\n"):
                if not ln.strip():
                    continue
                sc, dp, ct, ur = self._parse_tldr_line(ln)
                if ct:
                    display = f"{dp} {ct}" if dp else ct
                    parsed_lines.append((sc, display, ur))

        for signal_color, translated, urls in parsed_lines:
            row = tk.Frame(body, bg=BG_CARD)
            row.pack(fill="x", pady=2, anchor="w")

            # Signal dot (colored circle)
            if signal_color:
                tk.Label(row, text="\u25CF", fg=signal_color, bg=BG_CARD,
                         font=("Segoe UI", 12)).pack(side="left", padx=(0, 6), anchor="n")

            # Text + inline links container
            text_frame = tk.Frame(row, bg=BG_CARD)
            text_frame.pack(side="left", fill="x", expand=True)

            # Append <link> inline with the text
            display_text = translated
            if urls:
                display_text += "  "
            text_label = tk.Label(text_frame, text=display_text, fg=FG, bg=BG_CARD,
                                  font=FONT_UI, justify="left", anchor="w",
                                  wraplength=350)
            text_label.pack(side="left", fill="x", expand=True)

            def _rewrap_line(event, lbl=text_label):
                lbl.config(wraplength=max(event.width - 40, 100))
            text_frame.bind("<Configure>", _rewrap_line)

            # Clickable <link> inline, with tooltip on hover
            for url in urls:
                link_lbl = tk.Label(text_frame, text="<link>", fg=ACCENT, bg=BG_CARD,
                                    font=("Consolas", 8, "underline"), cursor="hand2")
                link_lbl.pack(side="left", padx=(0, 4))
                link_lbl.bind("<Button-1>", lambda e, u=url: webbrowser.open(u))
                # Tooltip
                tip = tk.Toplevel(link_lbl)
                tip.withdraw()
                tip.overrideredirect(True)
                tip_label = tk.Label(tip, text=url, fg="white", bg="#333333",
                                     font=("Consolas", 8), padx=6, pady=3)
                tip_label.pack()
                def _show_tip(event, t=tip):
                    t.geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
                    t.deiconify()
                def _hide_tip(event, t=tip):
                    t.withdraw()
                link_lbl.bind("<Enter>", _show_tip)
                link_lbl.bind("<Leave>", _hide_tip)

        # post_url (Discord) not bound to card — individual <link> labels handle URLs

    def _add_splits_card(self, reverse_splits: list[dict], split_status: list[dict]):
        """Splits card: reverse split history + pending/announced splits."""
        card = self._make_card(self.content_frame, title="Splits")
        body = tk.Frame(card, bg=BG_CARD, padx=10, pady=10)
        body.pack(fill="x")

        # Pending/Announced splits first (more urgent)
        if split_status:
            tk.Label(body, text="SPLIT STATUS", fg="#FFD600", bg=BG_CARD,
                     font=FONT_UI_BOLD, anchor="w").pack(fill="x", pady=(0, 4))
            for i, s in enumerate(split_status):
                row_bg = BG_ROW if i % 2 == 0 else BG_ROW_ALT
                row = tk.Frame(body, bg=row_bg,
                               highlightbackground=BORDER_INNER, highlightthickness=1)
                row.pack(fill="x", pady=2)
                inner = tk.Frame(row, bg=row_bg, padx=8, pady=6)
                inner.pack(fill="x")

                action = s.get("action_type") or "Unknown"
                split_from = s.get("split_from")
                split_to = s.get("split_to")
                eff_date = (s.get("effective_date") or "")[:10] if s.get("effective_date") else ""
                vote_date = (s.get("vote_date") or "")[:10] if s.get("vote_date") else ""
                filed = (s.get("filed_at") or "")[:10]
                doc_url = s.get("document_url") or ""

                # Action type label
                action_color = "#FFD600" if "Pending" in action else GREEN if "Approved" in action else FG
                tk.Label(inner, text=action, fg=action_color, bg=row_bg,
                         font=FONT_UI_BOLD, anchor="w").pack(fill="x")

                # Info line
                info = ""
                if split_from and split_to:
                    info += f"Ratio: {int(split_from)}:{int(split_to)}"
                if eff_date:
                    info += f"  |  Effective: {eff_date}"
                if vote_date:
                    info += f"  |  Vote: {vote_date}"
                if filed:
                    info += f"  |  Filed: {filed}"
                if info:
                    tk.Label(inner, text=info, fg=FG_DIM, bg=row_bg,
                             font=FONT_MONO, anchor="w").pack(fill="x", pady=(2, 0))

                # Details (pre-translated)
                details = s.get("_translated_details") or s.get("details") or ""
                if details:
                    det_label = tk.Label(inner, text=details, fg=FG, bg=row_bg,
                                         font=FONT_UI, anchor="w", wraplength=350,
                                         justify="left")
                    det_label.pack(fill="x", pady=(4, 0))
                    def _rewrap_det(event, lbl=det_label):
                        lbl.config(wraplength=max(event.width - 40, 100))
                    row.bind("<Configure>", _rewrap_det)

                if doc_url:
                    for w in (row, inner):
                        w.config(cursor="hand2")
                        w.bind("<Button-1>", lambda e, u=doc_url: webbrowser.open(u))

        # Reverse split history
        if reverse_splits:
            tk.Label(body, text="REVERSE SPLIT HISTORY", fg=RED, bg=BG_CARD,
                     font=FONT_UI_BOLD, anchor="w").pack(fill="x", pady=(8 if split_status else 0, 4))
            for i, rs in enumerate(reverse_splits):
                row_bg = BG_ROW if i % 2 == 0 else BG_ROW_ALT
                row = tk.Frame(body, bg=row_bg,
                               highlightbackground=BORDER_INNER, highlightthickness=1)
                row.pack(fill="x", pady=2)
                inner = tk.Frame(row, bg=row_bg, padx=8, pady=6)
                inner.pack(fill="x")

                date = (rs.get("execution_date") or rs.get("date") or rs.get("effective_date") or "")[:10]
                split_from = rs.get("split_from")
                split_to = rs.get("split_to")

                info = date if date else "Unknown date"
                if split_from and split_to:
                    info += f"  |  {int(split_from)}:{int(split_to)}"

                tk.Label(inner, text=info, fg=FG, bg=row_bg,
                         font=FONT_MONO, anchor="w").pack(fill="x")

    def _add_nasdaq_compliance_card(self, results: list[dict]):
        """NASDAQ Compliance card showing OPEN deficiencies with translated notes."""
        card = self._make_card(self.content_frame, title="NASDAQ Compliance")
        body = tk.Frame(card, bg=BG_CARD, padx=10, pady=10)
        body.pack(fill="x")

        for i, item in enumerate(results):
            row_bg = BG_ROW if i % 2 == 0 else BG_ROW_ALT
            row = tk.Frame(body, bg=row_bg,
                           highlightbackground=BORDER_INNER, highlightthickness=1)
            row.pack(fill="x", pady=2)
            inner = tk.Frame(row, bg=row_bg, padx=8, pady=6)
            inner.pack(fill="x")

            deficiency = item.get("deficiency") or "N/A"
            status = item.get("status") or ""
            market = item.get("market") or ""
            date = (item.get("date") or "")[:10]
            notes = item.get("notes") or ""

            # Deficiency label
            tk.Label(inner, text=deficiency, fg=RED, bg=row_bg,
                     font=FONT_UI_BOLD, anchor="w", wraplength=350,
                     justify="left").pack(fill="x")

            # Status + market + date
            detail = f"Status: {status}"
            if market:
                detail += f"  |  {market}"
            if date:
                detail += f"  |  {date}"
            tk.Label(inner, text=detail, fg=FG_DIM, bg=row_bg,
                     font=FONT_MONO, anchor="w").pack(fill="x", pady=(2, 0))

            # Notes (pre-translated in background)
            translated_notes = item.get("_translated_notes") or ""
            if translated_notes:
                notes_label = tk.Label(inner, text=translated_notes, fg="white", bg=row_bg,
                                       font=FONT_UI, anchor="w", wraplength=350,
                                       justify="left")
                notes_label.pack(fill="x", pady=(4, 0))
                def _rewrap_notes(event, lbl=notes_label):
                    lbl.config(wraplength=max(event.width - 40, 100))
                row.bind("<Configure>", _rewrap_notes)

    def _add_in_play_section(self, warrants: list[dict], convertibles: list[dict],
                             stock_price: float = 0.0, dilution_url: str = ""):
        card = self._make_card(self.content_frame, title="In Play Dilution")
        body = tk.Frame(card, bg=BG_CARD, padx=14, pady=10)
        body.pack(fill="x")

        if warrants:
            tk.Label(
                body, text="WARRANTS", fg="#FFD600", bg=BG_CARD,
                font=FONT_UI_BOLD, anchor="w",
            ).pack(fill="x", pady=(4, 4))
            for w in warrants:
                ex_price = w.get("warrants_exercise_price", 0) or 0
                in_money = stock_price > 0 and ex_price <= stock_price
                self._add_dilution_row(
                    body, w.get("details", ""),
                    f"Remaining: {fmt_millions(w.get('warrants_remaining'))}",
                    f"Strike: ${ex_price:.2f}",
                    (w.get("filed_at") or "")[:10],
                    in_money,
                )

        if convertibles:
            tk.Label(
                body, text="CONVERTIBLES", fg="#FFD600", bg=BG_CARD,
                font=FONT_UI_BOLD, anchor="w",
            ).pack(fill="x", pady=(8, 4))
            for c in convertibles:
                conv_price = c.get("conversion_price", 0) or 0
                in_money = stock_price > 0 and conv_price <= stock_price
                self._add_dilution_row(
                    body, c.get("details", ""),
                    f"Shares: {fmt_millions(c.get('underlying_shares_remaining'))}",
                    f"Conv: ${conv_price:.2f}",
                    (c.get("filed_at") or "")[:10],
                    in_money,
                )

        if dilution_url:
            self._bind_card_click(card, dilution_url)

    def _add_dilution_row(self, parent, details, remaining, price, filed,
                          price_above=False):
        # Green if strike/conv price <= stock price (in the money), orange otherwise
        highlight = "#4CAF50" if price_above else "#FF9800"

        row = tk.Frame(parent, bg=BG_ROW,
                       highlightbackground=BORDER_INNER, highlightthickness=1)
        row.pack(fill="x", pady=2)

        inner = tk.Frame(row, bg=BG_ROW, padx=10, pady=6)
        inner.pack(fill="x")

        # Line 1: details (truncated if long)
        detail_text = details if len(details) <= 60 else details[:57] + "..."
        tk.Label(inner, text=detail_text, fg="white", bg=BG_ROW,
                 font=FONT_UI, anchor="w").pack(fill="x")

        # Line 2: remaining | price | filed
        data_row = tk.Frame(inner, bg=BG_ROW)
        data_row.pack(fill="x", pady=(2, 0))
        tk.Label(data_row, text=remaining, fg=highlight, bg=BG_ROW,
                 font=FONT_MONO_BOLD).pack(side="left")
        tk.Label(data_row, text="  |  ", fg=FG_DIM, bg=BG_ROW,
                 font=FONT_MONO).pack(side="left")
        tk.Label(data_row, text=price, fg=highlight, bg=BG_ROW,
                 font=FONT_MONO_BOLD).pack(side="left")
        tk.Label(data_row, text=f"  |  Filed: {filed}", fg=FG_DIM, bg=BG_ROW,
                 font=FONT_MONO).pack(side="left")

    # ── Gainers panel ───────────────────────────────────────────────────────
    def _schedule_gainers_refresh(self):
        """Kick off the first gainers fetch."""
        self._trigger_gainers_refresh()

    def _trigger_gainers_refresh(self):
        """Fetch gainers in background thread."""
        self._gainers_status.config(text="loading...")

        def _fetch():
            gainers = fetch_top_gainers()
            self.root.after(0, self._update_gainers_ui, gainers)

        threading.Thread(target=_fetch, daemon=True).start()

    def _update_gainers_ui(self, gainers: list[dict]):
        """Rebuild the gainers list with fresh data."""
        self._gainers_data = gainers
        self._gainers_status.config(text=str(len(gainers)))

        # ── Detect newly appeared tickers ──
        new_ticker_set = {item["ticker"] for item in gainers}
        if self._prev_gainer_tickers:
            self._new_gainer_tickers = new_ticker_set - self._prev_gainer_tickers
            if self._new_gainer_tickers:
                import winsound
                try:
                    winsound.Beep(880, 120)
                    self.root.after(180, lambda: winsound.Beep(1100, 120))
                except Exception:
                    pass
                self._flash_gainers_header()
        else:
            self._new_gainer_tickers = set()
        self._prev_gainer_tickers = new_ticker_set

        # Clear existing rows
        for w in self._gainers_frame.winfo_children():
            w.destroy()

        if not gainers:
            tk.Label(self._gainers_frame, text="No gainers found",
                     fg=FG_DIM, bg=BG, font=FONT_UI).pack(pady=20)
        else:
            for item in gainers:
                self._build_gainer_row(item)

        # Schedule next refresh
        self.root.after(GAINERS_REFRESH_SECS * 1000, self._trigger_gainers_refresh)

    def _flash_gainers_header(self, step: int = 0):
        """Flash the gainers status label to signal new tickers appeared."""
        # Alternate: yellow bg with black text / back to normal
        if step % 2 == 0:
            self._gainers_status.config(bg="#FFD600", fg="black")
        else:
            self._gainers_status.config(bg=BG_CARD, fg=FG_DIM)
        if step < 4:
            self.root.after(220, self._flash_gainers_header, step + 1)
        else:
            self._gainers_status.config(bg=BG_CARD, fg=FG_DIM)

    def _build_gainer_row(self, item: dict):
        """Build a single clickable gainer row."""
        ticker = item.get("ticker", "")
        change_pct = item.get("todaysChangePerc", 0)
        price = item.get("price", 0) or 0
        volume = item.get("volume", 0) or 0

        is_selected = (ticker == self._selected_gainer)
        row_bg = BG_SELECTED if is_selected else BG_CARD
        border_color = BORDER_ACCENT if is_selected else BORDER

        row = tk.Frame(self._gainers_frame, bg=row_bg,
                       highlightbackground=border_color, highlightthickness=1,
                       cursor="hand2")
        row.pack(fill="x", padx=4, pady=2)

        inner = tk.Frame(row, bg=row_bg, padx=10, pady=6)
        inner.pack(fill="x")

        # Top line: ticker + risk badge + change %
        top = tk.Frame(inner, bg=row_bg)
        top.pack(fill="x")

        tk.Label(top, text=ticker, fg=ACCENT, bg=row_bg,
                 font=FONT_GAINER_TICKER, cursor="hand2").pack(side="left")

        # NEW badge for tickers that just appeared in the list
        if ticker in self._new_gainer_tickers:
            tk.Label(top, text=" NEW ", fg="black", bg="#FFD600",
                     font=("Consolas", 7, "bold"), padx=3, pady=1,
                     cursor="hand2").pack(side="left", padx=(4, 0))

        risk_level = item.get("_risk", "")
        if risk_level:
            tk.Label(top, text=f" {risk_level} ", fg="white",
                     bg=risk_bg(risk_level), font=("Consolas", 7, "bold"),
                     padx=4, pady=1, cursor="hand2").pack(side="left", padx=(4, 0))
        if item.get("_news_today"):
            tk.Label(top, text=" News ", fg="white", bg="#1F8FB3",
                     font=("Consolas", 7, "bold"), padx=4, pady=1,
                     cursor="hand2").pack(side="left", padx=(4, 0))
        pct_text = f"+{change_pct:.1f}%" if change_pct >= 0 else f"{change_pct:.1f}%"
        pct_color = GREEN if change_pct >= 0 else RED
        tk.Label(top, text=pct_text, fg=pct_color, bg=row_bg,
                 font=FONT_GAINER_PCT, cursor="hand2").pack(side="right")

        # Middle line: price + float badge + volume
        mid = tk.Frame(inner, bg=row_bg)
        mid.pack(fill="x")

        tk.Label(mid, text=fmt_price(price), fg=FG, bg=row_bg,
                 font=FONT_GAINER_DETAIL, cursor="hand2").pack(side="left")
        exchange = item.get("_exchange", "")
        if exchange:
            tk.Label(mid, text=f" {exchange}", fg=FG_DIM, bg=row_bg,
                     font=("Consolas", 7), cursor="hand2").pack(side="left")

        # Float badge: color-coded by size
        flt = item.get("_float")
        if flt:
            flt_m = flt / 1_000_000
            if flt_m < 5:
                flt_color = "#2F7D57"   # green — micro float
            elif flt_m < 25:
                flt_color = "#B9A816"   # yellow — small float
            else:
                flt_color = "#555555"   # gray — large float
            flt_label = fmt_millions(flt)
            tk.Label(mid, text=f" {flt_label} ", fg="white", bg=flt_color,
                     font=("Consolas", 7, "bold"), padx=3, pady=1,
                     cursor="hand2").pack(side="left", padx=(5, 0))

        tk.Label(mid, text=f"Vol {fmt_volume(volume)}", fg=FG_DIM, bg=row_bg,
                 font=FONT_GAINER_DETAIL, cursor="hand2").pack(side="right")

        # Bottom line: mcap / sector / country (float moved to mid line)
        mcap = item.get("_mcap")
        sector = item.get("_sector", "")
        country = item.get("_country", "")
        sector_short = {
            "Healthcare": "Health", "Technology": "Tech",
            "Industrials": "Indust", "Consumer Cyclical": "Cons Cyc",
            "Consumer Defensive": "Cons Def", "Communication Services": "Comms",
            "Financial Services": "Financ", "Basic Materials": "Materials",
            "Real Estate": "RE",
        }.get(sector, sector)
        info_parts = []
        if mcap:
            info_parts.append(fmt_millions(mcap))
        if sector_short:
            info_parts.append(sector_short)
        if country:
            info_parts.append(country)
        if info_parts:
            bot = tk.Frame(inner, bg=row_bg)
            bot.pack(fill="x")
            tk.Label(bot, text=" | ".join(info_parts), fg=FG_DIM, bg=row_bg,
                     font=FONT_GAINER_DETAIL, cursor="hand2").pack(side="left")
        else:
            bot = None

        # Bind click on all child widgets
        def on_click(e, t=ticker):
            self._on_gainer_click(t)

        click_targets = [row, inner, top, mid]
        if bot:
            click_targets.append(bot)
        for widget in click_targets:
            widget.bind("<Button-1>", on_click)
        for widget in (list(top.winfo_children()) + list(mid.winfo_children())
                       + (list(bot.winfo_children()) if bot else [])):
            widget.bind("<Button-1>", on_click)

    def _on_gainer_click(self, ticker: str):
        """Handle click on a gainer — select it and load Ask Edgar data."""
        self._selected_gainer = ticker
        # Rebuild gainers list to update selection highlight
        self._rebuild_gainers_list()
        # Load Ask Edgar data
        self._on_ticker_change(ticker)

    def _rebuild_gainers_list(self):
        """Rebuild gainer rows from cached data (updates selection state)."""
        for w in self._gainers_frame.winfo_children():
            w.destroy()
        for item in self._gainers_data:
            self._build_gainer_row(item)

    def _on_search(self, event=None):
        """Handle search box submit."""
        ticker = self.search_entry.get().strip().upper()
        if ticker:
            self.search_entry.delete(0, "end")
            self._selected_gainer = None
            self._rebuild_gainers_list()
            self._on_ticker_change(ticker)

    # ── Market Strength window ──
    # ── Chart window (TradingView widget via pywebview in subprocess) ──
    def _open_chart_window(self):
        """Open TradingView chart in the default browser for the current ticker."""
        ticker = self.current_ticker or "AAPL"
        url = f"https://www.tradingview.com/chart/?symbol={ticker}&interval=D"
        webbrowser.open(url)

    def _update_chart_ticker(self, ticker: str):
        """Update the chart ticker file so the chart process picks it up."""
        try:
            ticker_file = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), ".chart_ticker"
            )
            with open(ticker_file, "w") as f:
                f.write(ticker)
        except Exception:
            pass

    def _open_market_strength(self):
        """Open a new window showing market strength analysis."""
        from datetime import date as _date_cls

        win = tk.Toplevel(self.root)
        win.title("Market Strength")
        win.geometry("800x700")
        win.configure(bg=BG)
        win.attributes("-topmost", True)

        # Top bar: date input + fetch button
        top = tk.Frame(win, bg=BG_CARD, padx=10, pady=8)
        top.pack(fill="x")

        tk.Label(top, text="Date:", fg=FG_DIM, bg=BG_CARD,
                 font=FONT_UI_BOLD).pack(side="left", padx=(0, 6))

        date_entry = tk.Entry(top, bg=BG_ROW, fg=FG, insertbackground=FG,
                              font=FONT_UI_BOLD, width=12, relief="flat",
                              highlightbackground=BORDER, highlightthickness=1)
        date_entry.insert(0, _date_cls.today().isoformat())
        date_entry.pack(side="left", padx=(0, 6), ipady=3)

        fetch_btn = tk.Label(top, text="  Fetch  ", fg=BG, bg=ACCENT,
                             font=FONT_UI_BOLD, padx=8, pady=2, cursor="hand2")
        fetch_btn.pack(side="left")

        status_lbl = tk.Label(top, text="", fg=FG_DIM, bg=BG_CARD, font=FONT_UI)
        status_lbl.pack(side="left", padx=(10, 0))

        # Scrollable content area
        container = tk.Frame(win, bg=BG)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        content = tk.Frame(canvas, bg=BG)

        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_canvas_resize(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_resize)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        win.bind("<Destroy>", lambda e: self._main_canvas.bind_all("<MouseWheel>", self._main_mousewheel_fn) if e.widget == win else None)

        def _fetch_ms(event=None):
            date_str = date_entry.get().strip()
            if not date_str:
                return
            status_lbl.config(text="Loading...")
            for w in content.winfo_children():
                w.destroy()

            def _do_fetch():
                try:
                    resp = requests.get(
                        MARKET_STRENGTH_API_URL,
                        headers={"API-KEY": ASKEDGAR_API_KEY, "Content-Type": "application/json"},
                        params={"date": date_str},
                        timeout=15,
                    )
                    data = resp.json()
                    if data.get("status") == "success" and data.get("results"):
                        result = data["results"][0]
                        analysis = result.get("analysis", "")
                        performance = result.get("performance", "")
                        # Translate if enabled
                        t_analysis = translate_text(analysis) if analysis else ""
                        t_performance = translate_text(performance) if performance else ""
                        win.after(0, _show_results, date_str, t_analysis, t_performance)
                    else:
                        win.after(0, _show_error, f"No data for {date_str}")
                except Exception as e:
                    win.after(0, _show_error, str(e))

            threading.Thread(target=_do_fetch, daemon=True).start()

        def _show_error(msg):
            status_lbl.config(text="")
            tk.Label(content, text=msg, fg=RED, bg=BG, font=FONT_UI).pack(pady=20)

        def _replace_discord_emojis(text):
            replacements = {
                ":arrow_up:": "\u2B06",
                ":arrow_down:": "\u2B07",
                ":orange_circle:": "\U0001F7E0",
                ":green_circle:": "\U0001F7E2",
                ":red_circle:": "\U0001F534",
                ":heavy_minus_sign:": "\u2796",
                ":warning:": "\u26A0",
                ":white_check_mark:": "\u2705",
            }
            for k, v in replacements.items():
                text = text.replace(k, v)
            return text

        def _show_results(date_str, analysis, performance):
            status_lbl.config(text=f"Data for {date_str}")
            analysis = _replace_discord_emojis(analysis)
            performance = _replace_discord_emojis(performance)

            if analysis:
                # Analysis section
                card_a = tk.Frame(content, bg=BG_CARD,
                                  highlightbackground=BORDER, highlightthickness=1)
                card_a.pack(fill="x", padx=8, pady=(8, 0))
                tk.Label(card_a, text="Analysis", fg=ACCENT, bg=BG_CARD,
                         font=FONT_HEADER, anchor="w", padx=14, pady=10).pack(fill="x")
                tk.Frame(card_a, bg=BORDER, height=1).pack(fill="x")
                body_a = tk.Frame(card_a, bg=BG_CARD, padx=14, pady=10)
                body_a.pack(fill="x")
                lbl_a = tk.Label(body_a, text=analysis, fg=FG, bg=BG_CARD,
                                 font=FONT_UI, justify="left", anchor="nw",
                                 wraplength=700)
                lbl_a.pack(fill="x")
                def _rewrap_a(event, lbl=lbl_a):
                    lbl.config(wraplength=max(event.width - 30, 200))
                body_a.bind("<Configure>", _rewrap_a)

            if performance:
                # Performance section
                card_p = tk.Frame(content, bg=BG_CARD,
                                  highlightbackground=BORDER, highlightthickness=1)
                card_p.pack(fill="x", padx=8, pady=(8, 8))
                tk.Label(card_p, text="Performance", fg=ACCENT, bg=BG_CARD,
                         font=FONT_HEADER, anchor="w", padx=14, pady=10).pack(fill="x")
                tk.Frame(card_p, bg=BORDER, height=1).pack(fill="x")
                body_p = tk.Frame(card_p, bg=BG_CARD, padx=14, pady=10)
                body_p.pack(fill="x")
                lbl_p = tk.Label(body_p, text=performance, fg=FG, bg=BG_CARD,
                                 font=FONT_MONO, justify="left", anchor="nw",
                                 wraplength=700)
                lbl_p.pack(fill="x")
                def _rewrap_p(event, lbl=lbl_p):
                    lbl.config(wraplength=max(event.width - 30, 200))
                body_p.bind("<Configure>", _rewrap_p)

        fetch_btn.bind("<Button-1>", _fetch_ms)
        date_entry.bind("<Return>", _fetch_ms)

        # Auto-fetch on open
        _fetch_ms()

    # ── Dragging ──
    def _start_drag(self, event):
        self._drag_data["x"] = event.x
        self._drag_data["y"] = event.y

    def _on_drag(self, event):
        dx = event.x - self._drag_data["x"]
        dy = event.y - self._drag_data["y"]
        x = self.root.winfo_x() + dx
        y = self.root.winfo_y() + dy
        self.root.geometry(f"+{x}+{y}")

    # ── Monitor thread ──
    def _start_monitor(self):
        def poll():
            # ── One-time startup scan: log all window titles to help debug ──
            _startup_done = [False]

            while True:
                changed_ticker = None

                # ── DAS montage windows ──
                current = find_montage_windows()  # {hwnd: ticker}

                # On first poll, print all uppercase window titles for debugging
                if not _startup_done[0]:
                    _startup_done[0] = True
                    all_titles = []
                    def _log_all(hwnd, _):
                        if win32gui.IsWindowVisible(hwnd):
                            t = win32gui.GetWindowText(hwnd)
                            if t:
                                all_titles.append(t)
                    win32gui.EnumWindows(_log_all, None)
                    print(f"[DEBUG] Detected {len(current)} DAS montage windows: {list(current.values())}")
                    for t in all_titles:
                        if re.match(r'^[A-Z]{2,5}[\s\-]', t):
                            print(f"[DEBUG] Uppercase window: {repr(t)}")
                for hwnd, ticker in current.items():
                    old_ticker = self._known_windows.get(hwnd)
                    if old_ticker is not None and ticker != old_ticker:
                        changed_ticker = ticker
                        break
                if changed_ticker is None:
                    new_hwnds = set(current) - set(self._known_windows)
                    for hwnd in new_hwnds:
                        changed_ticker = current[hwnd]
                        break
                self._known_windows = current

                # ── thinkorswim chart windows ──
                if changed_ticker is None:
                    tos_current = find_tos_tickers()  # {hwnd: [tickers]}
                    for hwnd, tickers in tos_current.items():
                        old_tickers = self._known_tos.get(hwnd, [])
                        new_syms = [t for t in tickers if t not in old_tickers]
                        if new_syms:
                            changed_ticker = new_syms[0]
                            break
                    if changed_ticker is None:
                        new_hwnds = set(tos_current) - set(self._known_tos)
                        for hwnd in new_hwnds:
                            changed_ticker = tos_current[hwnd][0]
                            break
                    self._known_tos = tos_current

                if changed_ticker:
                    self.current_ticker = changed_ticker
                    self.root.after(0, self._on_ticker_change, changed_ticker)
                time.sleep(POLL_INTERVAL)

        threading.Thread(target=poll, daemon=True).start()

    def _on_ticker_change(self, ticker: str):
        self.current_ticker = ticker
        if self._selected_gainer and self._selected_gainer != ticker:
            self._selected_gainer = None
            self._rebuild_gainers_list()
        self._show_loading(ticker)
        self._update_chart_ticker(ticker)

        # ── Reutilizar datos del gainer si ya está en la lista ──────────────
        # Si el ticker aparece en _gainers_data, sus llamadas a screener,
        # dilution y news ya están en _api_cache (las hizo enrich() al cargar
        # los gainers). Precargamos también el objeto "screener sintético" con
        # los campos que sí tenemos del gainer para que fetch_screener_data()
        # no tenga que ir a la red en caso de que la caché haya expirado.
        gainer_match = next(
            (g for g in self._gainers_data if g.get("ticker", "").upper() == ticker.upper()),
            None
        )
        if gainer_match:
            print(f"[CACHE-HIT] {ticker} encontrado en Top Gainers — se reutilizan datos")
            now = time.time()
            # Screener sintético: reconstruimos el dict con los campos disponibles
            # para que fetch_screener_data() lo encuentre en caché
            if f"screener:{ticker}" not in _api_cache or \
               now - _api_cache[f"screener:{ticker}"][0] >= CACHE_TTL_SECS:
                synthetic_screener = {
                    "tradable_float": gainer_match.get("_float"),
                    "market_cap":     gainer_match.get("_mcap") or gainer_match.get("_tv_mcap"),
                    "sector":         gainer_match.get("_sector", ""),
                    "country":        gainer_match.get("_country", ""),
                    "price":          gainer_match.get("price", 0),
                }
                _api_cache[f"screener:{ticker}"] = (now, synthetic_screener)
                print(f"[CACHE-SEED] screener:{ticker} → desde gainer (float={gainer_match.get('_float')}, sector={gainer_match.get('_sector')})")
            # Dilution: si el gainer ya tiene _risk, construimos un dict mínimo
            if f"dilution:{ticker}" not in _api_cache or \
               now - _api_cache[f"dilution:{ticker}"][0] >= CACHE_TTL_SECS:
                risk = gainer_match.get("_risk", "")
                if risk:
                    synthetic_dilution = {"overall_offering_risk": risk}
                    _api_cache[f"dilution:{ticker}"] = (now, synthetic_dilution)
                    print(f"[CACHE-SEED] dilution:{ticker} → desde gainer (risk={risk})")
        else:
            print(f"[CACHE-MISS] {ticker} no está en Top Gainers — se consulta la API")

        def fetch():
            # ── Marcar si los datos base vienen del caché de gainers ──
            from_gainer_cache = gainer_match is not None
            with ThreadPoolExecutor(max_workers=13) as pool:
                f_dilution = pool.submit(fetch_dilution_data, ticker)
                f_screener = pool.submit(fetch_screener_data, ticker)
                f_news = pool.submit(fetch_news_and_grok, ticker)
                f_inplay = pool.submit(fetch_in_play_dilution, ticker)
                f_gap = pool.submit(fetch_gap_stats, ticker)
                f_offer = pool.submit(fetch_offerings, ticker)
                f_own = pool.submit(fetch_ownership, ticker)
                f_pump = pool.submit(fetch_pump_dump, ticker)
                f_tldr = pool.submit(fetch_research_tldr, ticker)
                f_rsplit = pool.submit(fetch_reverse_splits, ticker)
                f_sstatus = pool.submit(fetch_split_status, ticker)
                f_comp = pool.submit(fetch_nasdaq_compliance, ticker)
                f_chart = pool.submit(fetch_chart_analysis, ticker)

            dilution = f_dilution.result()
            screener = f_screener.result()
            news, grok_line, grok_date, grok_url, jmt415_notes = f_news.result()
            warrants, converts, stock_price = f_inplay.result()
            gap_stats = f_gap.result()
            recent_offerings = f_offer.result()
            ownership = f_own.result()
            pump_dump = f_pump.result()
            r_tldr = f_tldr.result()
            r_splits = f_rsplit.result()
            s_status = f_sstatus.result()
            compliance = f_comp.result()
            chart = f_chart.result()

            if from_gainer_cache:
                saved = []
                if screener:
                    saved.append("screener")
                if dilution:
                    saved.append("dilution")
                print(f"[CACHE-HIT] {ticker}: datos de gainers reutilizados para [{', '.join(saved)}]. "
                      f"APIs llamadas: inplay, gap, offerings, ownership, pump, tldr, rsplits, compliance, chart, news")

            # ── Claude LONG/SHORT signal (uses grok + news headlines) ──
            news_headlines = [extract_headline(n) for n in (news or [])]
            cname_for_signal = (screener.get("company_name") or screener.get("name") or "") if screener else ""
            claude_signal = None  # No auto-call — user triggers via ANALIZAR button

            # ── Pre-translate in parallel ──
            open_compliance = [c for c in (compliance or []) if (c.get("status") or "").upper() == "OPEN"]
            company_name = ""
            if SHOW_COMPANY_NAME and screener:
                cik = screener.get("cik") or ""
                company_name = fetch_company_name(cik)
            commentary = (dilution or {}).get("mgmt_commentary", "")

            translate_jobs = {}
            if news:
                for idx, ni in enumerate(news):
                    translate_jobs[f"news_{idx}"] = extract_headline(ni)
            if grok_line:
                translate_jobs["grok"] = grok_line
            if jmt415_notes:
                for idx, note in enumerate(jmt415_notes):
                    translate_jobs[f"jmt_{idx}"] = (note.get("summary") or note.get("title") or "Note").strip()
            if commentary:
                translate_jobs["commentary"] = commentary
            for idx, c in enumerate(open_compliance):
                n = c.get("notes") or ""
                if n:
                    translate_jobs[f"comp_{idx}"] = n
            # Filter split_status: hide executed and past pending votes
            from datetime import date as _date_cls
            today_str = _date_cls.today().isoformat()
            filtered_splits = []
            if s_status:
                for s in s_status:
                    if s.get("effective_date"):
                        continue  # already executed
                    action = (s.get("action_type") or "").lower()
                    if "pending" in action:
                        vd = (s.get("vote_date") or "")[:10]
                        if vd and vd < today_str:
                            continue  # vote already happened
                    filtered_splits.append(s)
            s_status = filtered_splits
            for idx, s in enumerate(s_status):
                det = s.get("details") or ""
                if det:
                    translate_jobs[f"split_{idx}"] = det
            tldr_parsed = []
            if r_tldr:
                tldr_item = r_tldr[0]
                tldr_text = (tldr_item.get("tldr_text") or tldr_item.get("analysis_text") or
                             tldr_item.get("summary") or tldr_item.get("report") or tldr_item.get("text") or "")
                for ln_idx, ln in enumerate(tldr_text.split("\n")):
                    if not ln.strip():
                        continue
                    _color, _date, _clean, _urls = self._parse_tldr_line(ln)
                    if _clean:
                        tldr_parsed.append((_color, _date, _urls, ln_idx))
                        translate_jobs[f"tldr_{ln_idx}"] = _clean

            # Fetch flag in parallel with translations
            country_code = screener.get("country", "") if screener else ""
            translated = {}
            flag_bytes = None
            with ThreadPoolExecutor(max_workers=12) as tpool:
                # Flag fetch
                f_flag = tpool.submit(fetch_flag_image, country_code)
                # Translations
                t_futures = {}
                if translate_jobs and TRANSLATE_ENABLED:
                    t_futures = {k: tpool.submit(translate_text, v) for k, v in translate_jobs.items()}
            for k, fut in t_futures.items():
                translated[k] = fut.result()
            flag_bytes = f_flag.result()

            if news:
                for idx, ni in enumerate(news):
                    ni["_translated_headline"] = translated.get(f"news_{idx}") or extract_headline(ni)
            if grok_line:
                grok_line = translated.get("grok") or grok_line
            if jmt415_notes:
                for idx, note in enumerate(jmt415_notes):
                    note["_translated_text"] = translated.get(f"jmt_{idx}") or (note.get("summary") or note.get("title") or "Note").strip()
            translated_commentary = translated.get("commentary") or ""
            for idx, c in enumerate(open_compliance):
                c["_translated_notes"] = translated.get(f"comp_{idx}") or ""
            for idx, s in enumerate(s_status):
                s["_translated_details"] = translated.get(f"split_{idx}") or s.get("details", "")
            if r_tldr and tldr_parsed:
                tldr_item = r_tldr[0]
                tlines = []
                for _color, _date, _urls, ln_idx in tldr_parsed:
                    t = translated.get(f"tldr_{ln_idx}") or ""
                    if _date:
                        t = f"{_date} {t}"
                    tlines.append((_color, t, _urls))
                tldr_item["_translated_lines"] = tlines

            history_rating = chart.get("rating", "") if chart else ""
            history_url = chart.get("post_url", "") if chart else ""

            # ── Overhead detection via Claude Vision (screenshot of DAS chart) ──
            has_overhead = False
            if CLAUDE_API_KEY:
                screenshot = capture_das_chart_window(ticker)
                if screenshot:
                    has_overhead = analyze_overhead_with_claude(ticker, screenshot)
                else:
                    print(f"[Overhead] No screenshot captured for {ticker}, skipping vision analysis")

            self.root.after(0, self._update_history_badge, history_rating, history_url)
            self.root.after(0, self._update_overhead_label, has_overhead)
            self.root.after(0, self._update_gainer_cache_badge, from_gainer_cache)
            self.root.after(0, self._show_data, ticker, dilution or {}, screener,
                            news, grok_line, grok_date, grok_url, warrants, converts, stock_price,
                            jmt415_notes, gap_stats, recent_offerings, ownership,
                            pump_dump, r_tldr, r_splits, s_status, open_compliance,
                            translated_commentary, company_name, flag_bytes, claude_signal)

        threading.Thread(target=fetch, daemon=True).start()

    def _on_analizar_click(self):
        """Trigger on-demand Claude analysis for the current ticker."""
        data = self._current_data_for_claude
        ticker = data.get("ticker", "")
        if not ticker:
            return
        if not CLAUDE_API_KEY:
            self._show_toast("Falta CLAUDE_API_KEY en .env", color="#A93232")
            return
        self._analizar_btn.config(text=" ... ", bg="#555555", fg="white")
        self._analizar_btn.unbind("<Button-1>")

        def _fetch():
            result = fetch_claude_signal(
                ticker,
                data.get("company_name", ""),
                data.get("grok_text", ""),
                data.get("news_headlines", []),
            )
            self.root.after(0, self._on_analizar_result, ticker, result, data)

        threading.Thread(target=_fetch, daemon=True).start()

    def _on_analizar_result(self, ticker: str, result: dict | None, original_data: dict):
        """Handle Claude analysis result: update card + save to history."""
        self._analizar_btn.config(text=" \u26a1 ANALIZAR ", bg="#FFD600", fg="black")
        self._analizar_btn.bind("<Button-1>", lambda e: self._on_analizar_click())

        if not result:
            self._show_toast("Error al llamar a Claude. Revisa tu API key.", color="#A93232")
            return

        import datetime
        entry = {
            "ticker": ticker,
            "timestamp": datetime.datetime.now().strftime("%d/%m/%y %H:%M"),
            "company_name": original_data.get("company_name", ""),
            "grok_text": original_data.get("grok_text", ""),
            "news_headlines": original_data.get("news_headlines", []),
            "signal": result.get("signal", ""),
            "summary": result.get("summary", ""),
            "catalysts": result.get("catalysts", []),
            "action": result.get("action", ""),
            "red_flags": result.get("red_flags", []),
        }
        self._claude_history.insert(0, entry)
        self._rebuild_claude_signal_card(ticker, result, original_data.get("company_name", ""))

    def _rebuild_claude_signal_card(self, ticker: str, claude_signal: dict, company_name: str):
        """Remove existing Claude signal card and inject a fresh one at the top."""
        for widget in self.content_frame.winfo_children():
            if getattr(widget, "_is_claude_card", False):
                widget.destroy()
                break

        signal    = claude_signal.get("signal", "")
        summary   = claude_signal.get("summary", "")
        catalysts = claude_signal.get("catalysts", [])
        action    = claude_signal.get("action", "")
        red_flags = claude_signal.get("red_flags", [])
        signal_color = "#2F7D57" if signal == "LONG" else "#A93232"
        dark_bg = "#0f1a14" if signal == "LONG" else "#1a0f0f"

        sig_card = tk.Frame(self.content_frame, bg=signal_color,
                            highlightbackground=signal_color, highlightthickness=2)
        sig_card._is_claude_card = True

        children = self.content_frame.winfo_children()
        if children:
            sig_card.pack(fill="x", padx=8, pady=(6, 0))
            sig_card.lift()
            sig_card.pack_forget()
            sig_card.pack(fill="x", padx=8, pady=(6, 0), before=children[0])
        else:
            sig_card.pack(fill="x", padx=8, pady=(6, 0))

        sig_inner = tk.Frame(sig_card, bg=dark_bg, padx=14, pady=12)
        sig_inner.pack(fill="x", padx=2, pady=(0, 2))

        sig_top = tk.Frame(sig_inner, bg=dark_bg)
        sig_top.pack(fill="x")
        tk.Label(sig_top, text=f"${ticker}", fg="white", bg=dark_bg,
                 font=("Segoe UI Semibold", 14)).pack(side="left")
        if company_name:
            tk.Label(sig_top, text=f"  {company_name}", fg=FG_DIM, bg=dark_bg,
                     font=FONT_UI).pack(side="left")
        tk.Label(sig_top, text=f"  {signal}  ", fg="white", bg=signal_color,
                 font=("Segoe UI Semibold", 13, "bold"), padx=10, pady=4).pack(side="right")

        if summary:
            tk.Frame(sig_inner, bg=BORDER, height=1).pack(fill="x", pady=(10, 6))
            tk.Label(sig_inner, text="RESUMEN", fg=FG_DIM, bg=dark_bg,
                     font=("Consolas", 8)).pack(anchor="w")
            lbl_sum = tk.Label(sig_inner, text=summary, fg=FG, bg=dark_bg,
                               font=FONT_UI, justify="left", anchor="w", wraplength=400)
            lbl_sum.pack(fill="x", pady=(4, 0))
            def _rw_sum(e, l=lbl_sum): l.config(wraplength=max(e.width - 30, 100))
            sig_inner.bind("<Configure>", _rw_sum)

        if catalysts:
            tk.Frame(sig_inner, bg=BORDER, height=1).pack(fill="x", pady=(10, 6))
            tk.Label(sig_inner, text="CATALIZADORES", fg=FG_DIM, bg=dark_bg,
                     font=("Consolas", 8)).pack(anchor="w")
            for cat in catalysts:
                row = tk.Frame(sig_inner, bg=dark_bg)
                row.pack(fill="x", pady=1)
                tk.Label(row, text="\u25C6", fg=ACCENT, bg=dark_bg,
                         font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))
                lbl_cat = tk.Label(row, text=cat, fg=ACCENT, bg=dark_bg,
                                   font=FONT_UI, anchor="w", wraplength=380, justify="left")
                lbl_cat.pack(side="left", fill="x", expand=True)

        if action:
            tk.Frame(sig_inner, bg=BORDER, height=1).pack(fill="x", pady=(10, 6))
            tk.Label(sig_inner, text="ACCION RECOMENDADA", fg=FG_DIM, bg=dark_bg,
                     font=("Consolas", 8)).pack(anchor="w")
            action_box = tk.Frame(sig_inner, bg=signal_color,
                                  highlightbackground=signal_color, highlightthickness=1)
            action_box.pack(fill="x", pady=(6, 0))
            action_inner = tk.Frame(action_box, bg="#1a1a1a", padx=10, pady=8)
            action_inner.pack(fill="x", padx=1, pady=(0, 1))
            top_act = tk.Frame(action_inner, bg="#1a1a1a")
            top_act.pack(fill="x")
            icon = "\U0001f4c9" if signal == "SHORT" else "\U0001f4c8"
            tk.Label(top_act, text=icon, bg="#1a1a1a", font=("Segoe UI", 11)).pack(side="left")
            tk.Label(top_act, text=f"  {signal}", fg=signal_color, bg="#1a1a1a",
                     font=("Segoe UI Semibold", 11, "bold")).pack(side="left")
            lbl_act = tk.Label(action_inner, text=action, fg=FG, bg="#1a1a1a",
                               font=FONT_UI, justify="left", anchor="w", wraplength=380)
            lbl_act.pack(fill="x", pady=(6, 0))
            def _rw_act(e, l=lbl_act): l.config(wraplength=max(e.width - 20, 100))
            action_inner.bind("<Configure>", _rw_act)

        if red_flags:
            tk.Frame(sig_inner, bg=BORDER, height=1).pack(fill="x", pady=(10, 6))
            tk.Label(sig_inner, text="RED FLAGS", fg=FG_DIM, bg=dark_bg,
                     font=("Consolas", 8)).pack(anchor="w")
            for flag in red_flags:
                row = tk.Frame(sig_inner, bg=dark_bg)
                row.pack(fill="x", pady=1)
                tk.Label(row, text="\U0001f6a9", bg=dark_bg,
                         font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))
                lbl_flag = tk.Label(row, text=flag, fg=RED, bg=dark_bg,
                                    font=FONT_UI, anchor="w", wraplength=380, justify="left")
                lbl_flag.pack(side="left", fill="x", expand=True)

        try:
            self.canvas.yview_moveto(0)
        except Exception:
            pass

    def _show_toast(self, message: str, color: str = "#2F7D57"):
        """Show a brief notification toast."""
        toast = tk.Toplevel(self.root)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        x = self.root.winfo_x() + self.root.winfo_width() // 2 - 180
        y = self.root.winfo_y() + 50
        toast.geometry(f"+{x}+{y}")
        tk.Label(toast, text=f"  {message}  ", fg="white", bg=color,
                 font=FONT_UI_BOLD, padx=16, pady=8).pack()
        toast.after(2500, toast.destroy)

    def _open_historial(self):
        """Open the Claude analysis history window."""
        win = tk.Toplevel(self.root)
        win.title("Historial Claude")
        win.geometry("700x600")
        win.configure(bg=BG)
        win.attributes("-topmost", True)

        top = tk.Frame(win, bg=BG_CARD, padx=12, pady=10)
        top.pack(fill="x")
        tk.Label(top, text="HISTORIAL DE ANÁLISIS CLAUDE", fg=ACCENT, bg=BG_CARD,
                 font=FONT_HEADER).pack(side="left")
        tk.Label(top, text=f"{len(self._claude_history)} consultas",
                 fg=FG_DIM, bg=BG_CARD, font=FONT_UI).pack(side="right")

        container = tk.Frame(win, bg=BG)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        content = tk.Frame(canvas, bg=BG)
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_resize(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_resize)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        win.bind("<Destroy>", lambda e: self._main_canvas.bind_all("<MouseWheel>", self._main_mousewheel_fn) if e.widget == win else None)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        if not self._claude_history:
            tk.Label(content, text="Sin consultas todavía.\nPulsa \u26a1 ANALIZAR para empezar.",
                     fg="#4A525C", bg=BG, font=("Segoe UI", 12),
                     justify="center").pack(pady=60)
        else:
            for entry in self._claude_history:
                signal = entry.get("signal", "")
                signal_color = "#2F7D57" if signal == "LONG" else "#A93232"
                dark_bg = "#0f1a14" if signal == "LONG" else "#1a0f0f"

                outer = tk.Frame(content, bg=signal_color,
                                 highlightbackground=signal_color, highlightthickness=2)
                outer.pack(fill="x", padx=10, pady=(8, 0))

                inner = tk.Frame(outer, bg=dark_bg, padx=14, pady=10)
                inner.pack(fill="x", padx=2, pady=(0, 2))

                hdr = tk.Frame(inner, bg=dark_bg)
                hdr.pack(fill="x")
                tk.Label(hdr, text=f"${entry['ticker']}", fg="white", bg=dark_bg,
                         font=("Segoe UI Semibold", 14)).pack(side="left")
                cname = entry.get("company_name", "")
                if cname:
                    tk.Label(hdr, text=f"  {cname}", fg=FG_DIM, bg=dark_bg,
                             font=FONT_UI).pack(side="left")
                tk.Label(hdr, text=entry.get("timestamp", ""), fg=FG_DIM, bg=dark_bg,
                         font=FONT_MONO).pack(side="right", padx=(0, 10))
                tk.Label(hdr, text=f"  {signal}  ", fg="white", bg=signal_color,
                         font=("Segoe UI Semibold", 12, "bold"), padx=8, pady=3).pack(side="right")

                summary = entry.get("summary", "")
                if summary:
                    tk.Frame(inner, bg=BORDER, height=1).pack(fill="x", pady=(8, 5))
                    lbl = tk.Label(inner, text=summary, fg=FG, bg=dark_bg,
                                   font=FONT_UI, justify="left", anchor="w", wraplength=620)
                    lbl.pack(fill="x")
                    def _rw(e, l=lbl): l.config(wraplength=max(e.width - 30, 200))
                    inner.bind("<Configure>", _rw)

                for cat in entry.get("catalysts", []):
                    r = tk.Frame(inner, bg=dark_bg)
                    r.pack(fill="x", pady=1)
                    tk.Label(r, text="\u25C6", fg=ACCENT, bg=dark_bg,
                             font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))
                    tk.Label(r, text=cat, fg=ACCENT, bg=dark_bg,
                             font=FONT_UI, anchor="w", wraplength=600).pack(side="left", fill="x", expand=True)

                action = entry.get("action", "")
                if action:
                    tk.Frame(inner, bg=BORDER, height=1).pack(fill="x", pady=(8, 4))
                    tk.Label(inner, text="ACCION", fg=FG_DIM, bg=dark_bg,
                             font=("Consolas", 8)).pack(anchor="w")
                    tk.Label(inner, text=action, fg=FG, bg=dark_bg,
                             font=FONT_UI, justify="left", anchor="w", wraplength=620).pack(fill="x", pady=(2, 0))

                for flag in entry.get("red_flags", []):
                    r = tk.Frame(inner, bg=dark_bg)
                    r.pack(fill="x", pady=1)
                    tk.Label(r, text="\U0001f6a9", bg=dark_bg,
                             font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))
                    tk.Label(r, text=flag, fg=RED, bg=dark_bg,
                             font=FONT_UI, anchor="w", wraplength=600).pack(side="left", fill="x", expand=True)

            tk.Frame(content, bg=BG, height=20).pack()

    def _open_recientes(self):
        """Open the recent tickers history window (last 15 tickers viewed)."""
        win = tk.Toplevel(self.root)
        win.title("Tickers Recientes")
        win.geometry("500x500")
        win.configure(bg=BG)
        win.attributes("-topmost", True)

        top = tk.Frame(win, bg=BG_CARD, padx=12, pady=10)
        top.pack(fill="x")
        tk.Label(top, text="TICKERS RECIENTES", fg=ACCENT, bg=BG_CARD,
                 font=FONT_HEADER).pack(side="left")
        tk.Label(top, text=f"{min(len(self._ticker_history), 15)} tickers",
                 fg=FG_DIM, bg=BG_CARD, font=FONT_UI).pack(side="right")

        container = tk.Frame(win, bg=BG)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        content = tk.Frame(canvas, bg=BG)
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_resize(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_resize)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        win.bind("<Destroy>", lambda e: self._main_canvas.bind_all(
            "<MouseWheel>", self._main_mousewheel_fn) if e.widget == win else None)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        if not self._ticker_history:
            tk.Label(content, text="Sin tickers todavía.\nCarga un ticker para empezar.",
                     fg="#4A525C", bg=BG, font=("Segoe UI", 12),
                     justify="center").pack(pady=60)
        else:
            for entry in self._ticker_history:
                risk = entry.get("risk", "")
                flt = entry.get("float")

                row_outer = tk.Frame(content, bg=BORDER, highlightthickness=0)
                row_outer.pack(fill="x", padx=10, pady=(6, 0))

                row = tk.Frame(row_outer, bg=BG_CARD, padx=14, pady=10, cursor="hand2")
                row.pack(fill="x", padx=1, pady=(0, 1))

                hdr = tk.Frame(row, bg=BG_CARD)
                hdr.pack(fill="x")

                tk.Label(hdr, text=entry["ticker"], fg=ACCENT, bg=BG_CARD,
                         font=("Segoe UI Semibold", 14)).pack(side="left")
                cname = entry.get("company_name", "")
                if cname:
                    tk.Label(hdr, text=f"  {cname}", fg=FG_DIM, bg=BG_CARD,
                             font=FONT_UI).pack(side="left")
                tk.Label(hdr, text=entry.get("timestamp", ""), fg=FG_DIM, bg=BG_CARD,
                         font=FONT_MONO).pack(side="right")

                # Risk + float badges
                badges = tk.Frame(row, bg=BG_CARD)
                badges.pack(fill="x", pady=(6, 0))
                if risk:
                    tk.Label(badges, text=f" {risk} ", fg="white",
                             bg=risk_bg(risk), font=("Consolas", 9, "bold"),
                             padx=6, pady=2).pack(side="left")
                if flt:
                    flt_m = flt / 1_000_000
                    flt_color = "#2F7D57" if flt_m < 5 else ("#B9A816" if flt_m < 25 else "#555555")
                    tk.Label(badges, text=f" Float: {fmt_millions(flt)} ", fg="white",
                             bg=flt_color, font=("Consolas", 9, "bold"),
                             padx=6, pady=2).pack(side="left", padx=(6, 0))
                if not risk and not flt:
                    tk.Label(badges, text="Sin datos de dilución", fg=FG_DIM, bg=BG_CARD,
                             font=FONT_UI).pack(side="left")

                # Click to reload ticker
                def _on_click(e, t=entry["ticker"]):
                    win.destroy()
                    self._on_ticker_change(t)
                for w in [row_outer, row, hdr, badges] + list(hdr.winfo_children()) + list(badges.winfo_children()):
                    try:
                        w.bind("<Button-1>", _on_click)
                    except Exception:
                        pass

            tk.Frame(content, bg=BG, height=20).pack()

    def _build_session_summary(self) -> dict:
        """Collect stats for the session summary."""
        from datetime import datetime as _dt
        now = _dt.now()
        duration = now - self._session_start
        total_secs = int(duration.total_seconds())
        hours, rem = divmod(total_secs, 3600)
        minutes, secs = divmod(rem, 60)
        duration_str = f"{hours}h {minutes}m {secs}s" if hours else f"{minutes}m {secs}s"

        tickers_viewed = list(dict.fromkeys(
            h["ticker"] for h in reversed(self._ticker_history)
        ))

        signals = self._claude_history
        longs  = [s for s in signals if s.get("signal") == "LONG"]
        shorts = [s for s in signals if s.get("signal") == "SHORT"]

        risk_counts = {}
        for h in self._ticker_history:
            r = h.get("risk", "")
            if r:
                risk_counts[r] = risk_counts.get(r, 0) + 1

        return {
            "start":         self._session_start.strftime("%H:%M:%S"),
            "end":           now.strftime("%H:%M:%S"),
            "duration":      duration_str,
            "tickers":       tickers_viewed,
            "total_tickers": len(tickers_viewed),
            "signals":       signals,
            "longs":         longs,
            "shorts":        shorts,
            "risk_counts":   risk_counts,
            "date":          now.strftime("%Y-%m-%d"),
        }

    def _save_session_log(self, summary: dict):
        """Write session summary to logs/session_YYYYMMDD_HHMMSS.txt"""
        try:
            from datetime import datetime as _dt
            logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
            os.makedirs(logs_dir, exist_ok=True)
            fname = f"session_{self._session_start.strftime('%Y%m%d_%H%M%S')}.txt"
            path = os.path.join(logs_dir, fname)
            lines = []
            lines.append("=" * 52)
            lines.append("  DILU-PIRATA  —  RESUMEN DE SESIÓN")
            lines.append("=" * 52)
            lines.append(f"  Fecha    : {summary['date']}")
            lines.append(f"  Inicio   : {summary['start']}")
            lines.append(f"  Cierre   : {summary['end']}")
            lines.append(f"  Duración : {summary['duration']}")
            lines.append("")
            lines.append(f"  Tickers analizados ({summary['total_tickers']}):")
            for t in summary["tickers"]:
                lines.append(f"    • {t}")
            if summary["risk_counts"]:
                lines.append("")
                lines.append("  Distribución de riesgo:")
                for level, count in sorted(summary["risk_counts"].items()):
                    lines.append(f"    {level:<10}: {count}")
            lines.append("")
            lines.append(f"  Señales Claude  : {len(summary['signals'])}")
            lines.append(f"    LONG          : {len(summary['longs'])}")
            lines.append(f"    SHORT         : {len(summary['shorts'])}")
            if summary["signals"]:
                lines.append("")
                lines.append("  Detalle señales:")
                for s in reversed(summary["signals"]):
                    sig   = s.get("signal", "?")
                    tick  = s.get("ticker", "?")
                    ts    = s.get("timestamp", "")
                    summ  = s.get("summary", "")
                    lines.append(f"    [{ts}] {sig:<6} ${tick}")
                    if summ:
                        lines.append(f"           {summ[:70]}")
            lines.append("")
            lines.append("=" * 52)
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
        except Exception as e:
            print(f"[Session log] Error saving: {e}")

    def _show_session_summary_window(self, summary: dict):
        """Show a styled session summary popup before closing."""
        win = tk.Toplevel(self.root)
        win.title("Resumen de Sesión")
        win.configure(bg=BG)
        win.attributes("-topmost", True)
        win.resizable(False, False)

        # Center on screen
        win.update_idletasks()
        w, h = 480, 520
        sx = self.root.winfo_x() + self.root.winfo_width() // 2 - w // 2
        sy = self.root.winfo_y() + self.root.winfo_height() // 2 - h // 2
        win.geometry(f"{w}x{h}+{sx}+{sy}")

        # Header
        hdr = tk.Frame(win, bg=ACCENT, padx=16, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="RESUMEN DE SESIÓN", fg=BG, bg=ACCENT,
                 font=("Segoe UI Semibold", 15)).pack(side="left")
        tk.Label(hdr, text=summary["date"], fg=BG, bg=ACCENT,
                 font=FONT_MONO).pack(side="right")

        # Scrollable body
        container = tk.Frame(win, bg=BG)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, bg=BG, highlightthickness=0)
        sb = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        body = tk.Frame(canvas, bg=BG)
        body.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def _row(parent, label, value, val_color=FG):
            f = tk.Frame(parent, bg=BG_CARD, padx=14, pady=6)
            f.pack(fill="x", padx=12, pady=1)
            tk.Label(f, text=label, fg=FG_DIM, bg=BG_CARD,
                     font=FONT_MONO, width=18, anchor="w").pack(side="left")
            tk.Label(f, text=value, fg=val_color, bg=BG_CARD,
                     font=FONT_UI_BOLD, anchor="w").pack(side="left")

        def _section(parent, title):
            tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", padx=12, pady=(10,2))
            tk.Label(parent, text=title, fg=ACCENT, bg=BG,
                     font=FONT_MONO_BOLD, anchor="w", padx=14).pack(fill="x")

        # ── Tiempo ──
        _section(body, "⏱  TIEMPO")
        _row(body, "Inicio",    summary["start"])
        _row(body, "Cierre",    summary["end"])
        _row(body, "Duración",  summary["duration"], ACCENT)

        # ── Tickers ──
        _section(body, f"📊  TICKERS  ({summary['total_tickers']})")
        if summary["tickers"]:
            chip_frame = tk.Frame(body, bg=BG, padx=14)
            chip_frame.pack(fill="x", pady=4)
            for i, t in enumerate(summary["tickers"]):
                h = self._ticker_history
                risk = next((x.get("risk","") for x in h if x["ticker"]==t), "")
                chip_bg = risk_bg(risk) if risk else "#2A3240"
                tk.Label(chip_frame, text=f" {t} ", fg="white", bg=chip_bg,
                         font=FONT_MONO_BOLD, padx=6, pady=2).grid(
                    row=i//5, column=i%5, padx=2, pady=2, sticky="w")
        else:
            tk.Label(body, text="  Sin tickers esta sesión", fg=FG_DIM, bg=BG,
                     font=FONT_UI).pack(anchor="w", padx=14)

        # ── Riesgo ──
        if summary["risk_counts"]:
            _section(body, "🔴  DISTRIBUCIÓN DE RIESGO")
            risk_frame = tk.Frame(body, bg=BG, padx=14)
            risk_frame.pack(fill="x", pady=4)
            for level, count in sorted(summary["risk_counts"].items()):
                tk.Label(risk_frame, text=f" {level}: {count} ",
                         fg="white", bg=risk_bg(level),
                         font=FONT_MONO_BOLD, padx=8, pady=3).pack(
                    side="left", padx=(0, 6))

        # ── Señales Claude ──
        _section(body, f"⚡  SEÑALES CLAUDE  ({len(summary['signals'])})")
        sig_frame = tk.Frame(body, bg=BG, padx=14)
        sig_frame.pack(fill="x", pady=4)
        tk.Label(sig_frame, text=f" LONG: {len(summary['longs'])} ",
                 fg="white", bg="#2F7D57",
                 font=FONT_MONO_BOLD, padx=10, pady=3).pack(side="left")
        tk.Label(sig_frame, text=f" SHORT: {len(summary['shorts'])} ",
                 fg="white", bg="#A93232",
                 font=FONT_MONO_BOLD, padx=10, pady=3).pack(side="left", padx=(6, 0))

        if summary["signals"]:
            for s in reversed(summary["signals"]):
                sig  = s.get("signal", "?")
                tick = s.get("ticker", "?")
                ts   = s.get("timestamp", "")
                summ = s.get("summary", "")
                sig_bg  = "#2F7D57" if sig == "LONG" else "#A93232"
                dark_bg = "#0f1a14" if sig == "LONG" else "#1a0f0f"
                row = tk.Frame(body, bg=dark_bg,
                               highlightbackground=sig_bg, highlightthickness=1)
                row.pack(fill="x", padx=12, pady=2)
                inner = tk.Frame(row, bg=dark_bg, padx=10, pady=6)
                inner.pack(fill="x")
                top_r = tk.Frame(inner, bg=dark_bg)
                top_r.pack(fill="x")
                tk.Label(top_r, text=f"${tick}", fg="white", bg=dark_bg,
                         font=FONT_UI_BOLD).pack(side="left")
                tk.Label(top_r, text=f" {sig} ", fg="white", bg=sig_bg,
                         font=FONT_MONO_BOLD, padx=6, pady=1).pack(side="right")
                tk.Label(top_r, text=ts, fg=FG_DIM, bg=dark_bg,
                         font=FONT_MONO).pack(side="right", padx=(0,6))
                if summ:
                    lbl = tk.Label(inner, text=summ, fg=FG_DIM, bg=dark_bg,
                                   font=FONT_UI, anchor="w", wraplength=400,
                                   justify="left")
                    lbl.pack(fill="x", pady=(3, 0))

        # ── Footer buttons ──
        tk.Frame(win, bg=BORDER, height=1).pack(fill="x")
        btn_frame = tk.Frame(win, bg=BG_CARD, padx=14, pady=10)
        btn_frame.pack(fill="x")

        tk.Label(btn_frame, text="Log guardado en /logs/", fg=FG_DIM, bg=BG_CARD,
                 font=FONT_MONO).pack(side="left")

        close_btn = tk.Label(btn_frame, text="  CERRAR  ", fg=BG, bg=RED,
                             font=FONT_UI_BOLD, padx=10, pady=4, cursor="hand2")
        close_btn.pack(side="right")
        close_btn.bind("<Button-1>", lambda e: self.root.destroy())

        # Closing the popup also closes the app
        win.protocol("WM_DELETE_WINDOW", self.root.destroy)

    def _on_close(self):
        """Show session summary, save log, close chart window, then exit."""
        try:
            close_file = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), ".chart_close"
            )
            with open(close_file, "w") as f:
                f.write("close")
        except Exception:
            pass

        summary = self._build_session_summary()
        self._save_session_log(summary)

        # Only show popup if at least 1 ticker was viewed
        if summary["total_tickers"] > 0:
            self._show_session_summary_window(summary)
        else:
            self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()


if __name__ == "__main__":
    app = DilutionOverlay()
    app.run()

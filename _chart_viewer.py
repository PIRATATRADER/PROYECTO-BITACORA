"""
TradingView Chart Viewer — runs as a separate process from the main app.
Loads the full TradingView chart via iframe using the session cookie.
Reads ticker from .chart_ticker file and updates the chart automatically.
"""
import os
import json
import time
import webview

try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
except ImportError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TICKER_FILE = os.path.join(BASE_DIR, ".chart_ticker")
POSITION_FILE = os.path.join(BASE_DIR, "chart_position.cfg")
SAVE_REQUEST_FILE = os.path.join(BASE_DIR, ".chart_save_position")
CLOSE_REQUEST_FILE = os.path.join(BASE_DIR, ".chart_close")


def get_chart_url(symbol="AAPL"):
    """TradingView full chart URL with daily interval."""
    return f"https://www.tradingview.com/chart/?symbol={symbol}&interval=D"


def read_ticker():
    """Read current ticker from shared file."""
    try:
        if os.path.exists(TICKER_FILE):
            with open(TICKER_FILE, "r") as f:
                return f.read().strip().upper()
    except Exception:
        pass
    return ""


def load_position():
    """Load saved chart window position."""
    try:
        if os.path.exists(POSITION_FILE):
            with open(POSITION_FILE, "r") as f:
                return json.load(f)
    except Exception:
        pass
    return {"x": 100, "y": 100, "width": 1000, "height": 650}


def save_position(window):
    """Save current chart window position to file."""
    try:
        pos = {
            "x": window.x,
            "y": window.y,
            "width": window.width,
            "height": window.height,
        }
        with open(POSITION_FILE, "w") as f:
            json.dump(pos, f)
    except Exception:
        pass


def watch_ticker(window):
    """Poll the ticker file and update chart when ticker changes."""
    last_ticker = read_ticker()
    while True:
        time.sleep(1)
        try:
            # Check for ticker change
            current = read_ticker()
            if current and current != last_ticker:
                last_ticker = current
                window.load_url(get_chart_url(current))

            # Check for save position request
            if os.path.exists(SAVE_REQUEST_FILE):
                os.remove(SAVE_REQUEST_FILE)
                save_position(window)

            # Check for close request
            if os.path.exists(CLOSE_REQUEST_FILE):
                os.remove(CLOSE_REQUEST_FILE)
                window.destroy()
                break
        except Exception:
            pass


def main():
    # Clean up stale signals from previous session
    for f in (SAVE_REQUEST_FILE, CLOSE_REQUEST_FILE):
        if os.path.exists(f):
            os.remove(f)

    ticker = read_ticker() or "AAPL"
    pos = load_position()

    # Inject TradingView session cookie for full chart access
    session_id = os.environ.get("TRADINGVIEW_SESSION_ID", "")

    # Load TradingView homepage first to set cookie on the correct domain
    initial_url = "https://www.tradingview.com/robots.txt" if session_id else get_chart_url(ticker)

    window = webview.create_window(
        "TradingView Chart",
        url=initial_url,
        x=pos["x"],
        y=pos["y"],
        width=pos["width"],
        height=pos["height"],
    )

    def start_watcher(window):
        """Set cookie on TradingView domain, then load chart."""
        time.sleep(1)
        if session_id:
            window.evaluate_js(f'document.cookie="sessionid={session_id};domain=.tradingview.com;path=/;max-age=31536000";')
            time.sleep(0.5)
        window.load_url(get_chart_url(ticker))
        watch_ticker(window)

    webview.start(func=start_watcher, args=(window,))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Chart Viewer Error", str(e))
        root.destroy()

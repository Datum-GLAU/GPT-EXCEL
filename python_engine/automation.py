import schedule
import time
import threading
import os
from analysis import analyze_data
from charts import create_chart
from SQLite_storage import log_automation

WATCHED_FILE = os.environ.get("GPT_EXCEL_WATCH_FILE", "uploads/latest.xlsx")
AUTOMATION_ENABLED = os.environ.get("GPT_EXCEL_AUTOMATION", "false").lower() == "true"


def auto_analyze():
    print("[Automation] Running scheduled analysis...")
    if not os.path.exists(WATCHED_FILE):
        log_automation("analyze", "skipped", f"Watched file not found: {WATCHED_FILE}")
        return
    try:
        result = analyze_data(WATCHED_FILE)
        msg = f"Rows: {result['total_rows']}, Cols: {result['total_columns']}"
        log_automation("analyze", "success", msg)
        print(f"[Automation] Analysis done — {msg}")
    except Exception as e:
        log_automation("analyze", "error", str(e))
        print(f"[Automation] Analysis failed: {e}")


def auto_chart():
    print("[Automation] Running scheduled chart generation...")
    if not os.path.exists(WATCHED_FILE):
        log_automation("chart", "skipped", f"Watched file not found: {WATCHED_FILE}")
        return
    try:
        result = create_chart(WATCHED_FILE, chart_type="auto")
        log_automation("chart", "success", result)
        print(f"[Automation] {result}")
    except Exception as e:
        log_automation("chart", "error", str(e))
        print(f"[Automation] Chart failed: {e}")


def start_scheduler():
    schedule.every(5).minutes.do(auto_analyze)
    schedule.every(10).minutes.do(auto_chart)
    log_automation("scheduler", "started", "Scheduler initialized")
    print("[Automation] Scheduler started.")
    while True:
        schedule.run_pending()
        time.sleep(1)


def run_in_background():
    if not AUTOMATION_ENABLED:
        print("[Automation] Background scheduler disabled.")
        return
    thread = threading.Thread(target=start_scheduler, daemon=True)
    thread.start()
    print("[Automation] Background scheduler running.")

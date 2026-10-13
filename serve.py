from waitress import serve
from app import application, _startup
from cbsl import ensure_directories_exist, kill_excel_instances
import scheduler as _sched

if __name__ == "__main__":
    _startup()
    kill_excel_instances()
    ensure_directories_exist()

    # Start scheduler if not already running (ifrs.py starts it at import time;
    # this guard prevents a double-start when both run together)
    if not (_sched._scheduler and _sched._scheduler.running):
        _sched.start_scheduler()
        print("Scheduler started.")

    print("Starting Waitress server on http://0.0.0.0:5000")
    serve(application, host="0.0.0.0", port=5000)
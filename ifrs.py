import os
import subprocess
import sys
import threading
import time
from datetime import date
from pathlib import Path

from flask import Flask, flash, jsonify, redirect, render_template, request, url_for

app = Flask(__name__, template_folder="templates")
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-me-in-prod")

BASE_DIR = Path(__file__).resolve().parent

# ── IFRS reports definition ────────────────────────────────────────────────
IFRS_REPORTS = [
    {"id": "IFRS-01", "num": "01", "name": "Portfolio 4 ECL",          "desc": "Expected Credit Loss – Portfolio 4",        "script": "ifrs_01_portfolio4_ecl.py"},
    {"id": "IFRS-02", "num": "02", "name": "ECL Model",                "desc": "ECL model computation",                     "script": "ifrs_02_ecl_model.py"},
    {"id": "IFRS-03", "num": "03", "name": "Summary",                  "desc": "IFRS 9 ECL summary",                        "script": "ifrs_03_summary.py"},
    {"id": "IFRS-04", "num": "04", "name": "PD Data Working",          "desc": "Probability of Default – data working",     "script": "ifrs_04_pd_data_working.py"},
    {"id": "IFRS-05", "num": "05", "name": "Historic PD Calculation",  "desc": "Historical PD rate calculations",           "script": "ifrs_05_historic_pd_calc.py"},
    {"id": "IFRS-06", "num": "06", "name": "First Default Date Working","desc": "First default date derivation",             "script": "ifrs_06_first_default_date.py"},
    {"id": "IFRS-07", "num": "07", "name": "Write Backs Working",      "desc": "Write-back adjustments working",            "script": "ifrs_07_write_backs.py"},
    {"id": "IFRS-08", "num": "08", "name": "LGD Working",              "desc": "Loss Given Default working",                "script": "ifrs_08_lgd_working.py"},
    {"id": "IFRS-09", "num": "09", "name": "PD Loan",                  "desc": "Probability of Default – Loans",            "script": "ifrs_09_pd_loan.py"},
    {"id": "IFRS-10", "num": "10", "name": "PD Lease",                 "desc": "Probability of Default – Leases",           "script": "ifrs_10_pd_lease.py"},
]

_REPORT_BY_ID = {r["id"]: r for r in IFRS_REPORTS}

# ── Shared run state ───────────────────────────────────────────────────────
_lock = threading.Lock()
_state = {
    "stage":          "idle",   # idle | running | done | error | stopping
    "running":        False,
    "current_report": "",
    "message":        "",
    "log_lines":      [],
    "checklist":      [],       # [{"id", "label", "status", "elapsed"}]
    "stop_requested": False,
    "current_proc":   None,
}


def _reset_state():
    with _lock:
        _state.update({
            "stage":          "idle",
            "running":        False,
            "current_report": "",
            "message":        "",
            "log_lines":      [],
            "checklist":      [],
            "stop_requested": False,
            "current_proc":   None,
        })


def _log(line: str):
    with _lock:
        _state["log_lines"].append(line)


def _set_checklist_status(report_id: str, status: str, elapsed: float | None = None):
    with _lock:
        for item in _state["checklist"]:
            if item["id"] == report_id:
                item["status"] = status
                if elapsed is not None:
                    item["elapsed"] = round(elapsed, 1)
                break


# ── Background worker ──────────────────────────────────────────────────────

def _run_worker(report: dict, run_date: str):
    """Run a single IFRS report script in a background thread."""
    script_path = BASE_DIR / "report_automations" / report["script"]

    with _lock:
        _state["stage"]          = "running"
        _state["running"]        = True
        _state["current_report"] = report["name"]
        _state["message"]        = f"Running {report['name']}…"
        _state["checklist"]      = [{
            "id":      report["id"],
            "label":   f"{report['num']}. {report['name']}",
            "status":  "running",
            "elapsed": None,
        }]

    _log(f"{'=' * 60}")
    _log(f"[IFRS] Starting: {report['name']}")
    _log(f"[IFRS] Date: {run_date}")
    _log(f"{'=' * 60}")

    start = time.time()
    success = True

    if not script_path.exists():
        _log(f"[WARN] Script not found: {script_path.name} — skipping.")
        _set_checklist_status(report["id"], "failed")
        success = False
    else:
        try:
            proc = subprocess.Popen(
                [sys.executable, str(script_path), run_date],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                cwd=str(BASE_DIR),
            )
            with _lock:
                _state["current_proc"] = proc

            for line in proc.stdout:
                with _lock:
                    if _state["stop_requested"]:
                        proc.terminate()
                        break
                _log(line.rstrip())

            proc.wait()

            with _lock:
                _state["current_proc"] = None
                stopped = _state["stop_requested"]

            elapsed = time.time() - start
            if stopped:
                _set_checklist_status(report["id"], "stopped", elapsed)
                _log(f"[WARN] Run stopped by user after {elapsed:.1f}s.")
                success = False
            elif proc.returncode != 0:
                _set_checklist_status(report["id"], "failed", elapsed)
                _log(f"[ERROR] Script exited with code {proc.returncode}.")
                success = False
            else:
                _set_checklist_status(report["id"], "done", elapsed)
                _log(f"[DONE] {report['name']} completed in {elapsed:.1f}s.")

        except Exception as exc:
            _set_checklist_status(report["id"], "failed")
            _log(f"[ERROR] {exc}")
            success = False

    _log(f"{'=' * 60}")

    with _lock:
        _state["running"]        = False
        _state["current_report"] = ""
        _state["stage"]          = "done" if success else "error"
        _state["message"]        = "Completed successfully." if success else "Finished with errors."


# ── Routes ─────────────────────────────────────────────────────────────────

@app.route("/")
def ifrs_dashboard():
    today_iso = date.today().isoformat()
    last_runs: dict = {}  # TODO: persist last-run dates per report if needed
    return render_template("ifrs_dashboard.html",
                           reports=IFRS_REPORTS,
                           today_iso=today_iso,
                           last_runs=last_runs)


@app.route("/run-report", methods=["POST"])
def run_report():
    report_id = request.form.get("report_id", "").strip()
    run_date  = request.form.get("date", "").strip()

    if not report_id or report_id not in _REPORT_BY_ID:
        flash("Invalid report selected.", "error")
        return redirect(url_for("ifrs_dashboard"))

    if not run_date:
        flash("Please select a date.", "warning")
        return redirect(url_for("ifrs_dashboard"))

    # Convert yyyy-mm-dd → MM/DD/YYYY for scripts
    try:
        y, m, d = run_date.split("-")
        script_date = f"{m}/{d}/{y}"
    except ValueError:
        flash("Invalid date format.", "error")
        return redirect(url_for("ifrs_dashboard"))

    with _lock:
        if _state["running"]:
            flash("A report is already running. Please wait.", "warning")
            return redirect(url_for("ifrs_dashboard"))

    _reset_state()
    report = _REPORT_BY_ID[report_id]
    t = threading.Thread(target=_run_worker, args=(report, script_date), daemon=True)
    t.start()

    return redirect(url_for("ifrs_dashboard"))


@app.route("/stop", methods=["POST"])
def stop():
    with _lock:
        _state["stop_requested"] = True
        _state["stage"]          = "stopping"
        proc = _state.get("current_proc")
    if proc:
        try:
            proc.terminate()
        except Exception:
            pass
    return jsonify({"ok": True})


@app.route("/reset", methods=["POST"])
def reset():
    _reset_state()
    return jsonify({"ok": True})


@app.route("/status")
def status():
    with _lock:
        return jsonify({
            "stage":          _state["stage"],
            "running":        _state["running"],
            "current_report": _state["current_report"],
            "message":        _state["message"],
        })


@app.route("/logs")
def logs():
    with _lock:
        lines = list(_state["log_lines"])
    return "\n".join(lines), 200, {"Content-Type": "text/plain; charset=utf-8"}


@app.route("/checklist")
def checklist():
    with _lock:
        items = list(_state["checklist"])
    return jsonify(items)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5002, debug=True)

Report Automation Web App

Features
- Create dated output folders under `outputs/[frequency]/YYYY-MM-DD`.
- Create subfolders inside the dated folder for each Python file in `report_automations/`.
- Mark completion dates per frequency; appends to `outputs/[frequency]/completed_dates.txt`.
- Start automation for weekly: finds the latest folder matching the last completed date and copies it to `working/weekly`.

Prerequisites
- Python 3.10+
- Windows PowerShell or Command Prompt

Setup
```
cd C:\CBSL\Script
python -m venv .venv
. .venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Run
```
python app.py
```
Then open `http://127.0.0.1:5000` in your browser.

Notes
- Dates are stored and foldered as `YYYY-MM-DD` to avoid ambiguity.
- Weekly automation copies to `working/weekly`, replacing any existing folder.
- Report subfolder names are derived from `report_automations/*.py` (file stem names).


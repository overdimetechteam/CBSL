import os
import shutil

# ── Configuration ──────────────────────────────────────────────────────────────
SOURCE_DIR        = r"Z:\ALCLREPORTS\MATINT"
SUMMARY_DIR       = r"C:\CBSL\Script\outputs\Summary"
COMPLETED_DATES   = os.path.join(SUMMARY_DIR, "completed_dates.txt")
# ───────────────────────────────────────────────────────────────────────────────


def get_last_date(filepath: str) -> str:
    with open(filepath, "r") as f:
        lines = [l.strip() for l in f.readlines() if l.strip()]
    if not lines:
        raise ValueError(f"No dates found in {filepath}")
    return lines[-1]


def get_latest_file(folder: str) -> str:
    files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if os.path.isfile(os.path.join(folder, f))
    ]
    if not files:
        raise FileNotFoundError(f"No files found in source folder: {folder}")
    return max(files, key=os.path.getmtime)


def copy_latest_file(source_dir: str, dest_dir: str) -> None:
    if not os.path.exists(source_dir):
        print(f"[ERROR] Source folder not found: {source_dir}")
        return

    latest = get_latest_file(source_dir)
    filename = os.path.basename(latest)
    dest_path = os.path.join(dest_dir, filename)

    os.makedirs(dest_dir, exist_ok=True)
    print(f"[INFO] Source file  : {latest}")
    print(f"[INFO] Destination  : {dest_path}")

    shutil.copy2(latest, dest_path)
    print(f"[OK]   Copied (replaced if existed): {filename}")


def main():
    if not os.path.exists(COMPLETED_DATES):
        print(f"[ERROR] completed_dates.txt not found: {COMPLETED_DATES}")
    else:
        date_folder = get_last_date(COMPLETED_DATES)
        dest_dir = os.path.join(SUMMARY_DIR, date_folder, "Input")
        print(f"[INFO] Date folder  : {date_folder}")
        copy_latest_file(SOURCE_DIR, dest_dir)
        return True

if __name__ == "__main__":
    main()

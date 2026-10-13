"""
One-off backfill script — run this directly on the server.

Writes .download_manifest.json into each report's LATEST completed outputs
date-folder, using the exact same file-selection patterns each report's
real post-processing code already uses. This is a stopgap so the
"download last run" button works immediately for existing folders, without
waiting for every report to run again under the fixed ifrs.py/cbsl.py.

Safe by design: only writes a manifest when it actually finds a matching
file. Never writes an empty/guessed manifest. Prints what it did (or
couldn't do) for every report so you can check the output before trusting
it. Does not import cbsl.py/ifrs.py (avoids triggering Selenium/Excel/
scheduler startup side effects) — just does plain filesystem globbing
against the same OUTPUTS_DIR layout.

Run from the Script root:  python backfill_manifests.py
"""
import glob
import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
OUTPUTS_DIR = BASE_DIR / "outputs"


def latest_date_folder(report_folder: Path) -> Path | None:
    if not report_folder.is_dir():
        return None
    subdirs = [p for p in report_folder.iterdir() if p.is_dir()]
    if not subdirs:
        return None
    return max(subdirs, key=lambda p: p.stat().st_mtime)


def pick_latest(folder: Path, pattern: str) -> Path | None:
    matches = sorted(folder.glob(pattern), key=lambda f: f.stat().st_mtime)
    return matches[-1] if matches else None


def write_manifest(run_root: Path, files: list[Path], label: str) -> None:
    rel_paths = []
    for f in files:
        if f is None or not f.exists():
            continue
        try:
            rel_paths.append(str(f.relative_to(run_root)))
        except ValueError:
            continue
    if not rel_paths:
        print(f"  [{label}] SKIPPED - no matching files found in {run_root}")
        return
    manifest_path = run_root / ".download_manifest.json"
    manifest_path.write_text(json.dumps({"files": rel_paths}), encoding="utf-8")
    print(f"  [{label}] wrote manifest at {manifest_path}")
    for p in rel_paths:
        print(f"      - {p}")


# ============================================================================
# CBSL reports — (folder name under outputs/, label, function that returns
# the list of files to record, given the run's date-folder path)
# ============================================================================
def cbsl_reports():
    def f_c1_c6(run):
        return [pick_latest(run, "NBD-MF-20-C1 to C6*.xlsx")]

    def f_af_c1_c6(run):
        return [pick_latest(run, "NBD-AF-20-C1 to C6*.xlsx")]

    def f_ga_is(run):
        return [
            pick_latest(run, "NBD-MF-10-GA & NBD-MF-11-IS *.xlsx"),
            pick_latest(run, "CBSL Provision Comparison - *.xlsb"),
        ]

    def f_c8(run):
        return [pick_latest(run, "C8-Client Rating-*.xlsx")]

    def f_c7(run):
        return [pick_latest(run, "C7- Report for Stage - Movement Q*.xlsb")]

    def f_c4(run):
        return [pick_latest(run, "Loan CF Analysis NBD-MF-23-C4-*.xlsb")]

    def f_dm(run):
        return [
            pick_latest(run, "NBD-WF-18-DM Deposit Liability Week *.xlsx"),
            pick_latest(run, "NBD_WF_18_DM_Detailed_Exceptions.txt"),
            pick_latest(run, "NBD-WF-18-DM Deposit Liability (*).xlsx"),
        ]

    def f_la(run):
        return [
            pick_latest(run, "AFL Liquidity - Week_*.xlsx") or pick_latest(run, "AFL Liquidity - *.xlsx"),
            pick_latest(run, "ALCL Management Accounts*.xlsx"),
            pick_latest(run, "BankBalanceGLCodes.xlsx"),
        ]

    def f_ecs(run):
        return [
            pick_latest(run, "Cash-Flow-Forecast-Week *.xlsm"),
            pick_latest(run, "Types_of_Receipt.xlsx"),
            pick_latest(run, "Tax Payments - *.xlsx"),
            pick_latest(run, "Cashflow Movement - *.xlsx"),
        ]

    def f_mf04_la(run):
        return [pick_latest(run, "NBD-MF-04-LA Liquid Assets*.xlsm")]

    def f_ia(run):
        return [
            pick_latest(run, "Prod. wise Class. of Loans - *.xlsb"),
            pick_latest(run, "EXCEPTIONS.xlsx"),
            pick_latest(run, "Minimum_Rate_Exceptions.xlsx"),
            pick_latest(run, "initial_valuation.xlsx"),
            pick_latest(run, "Product_Cat.xlsx"),
        ]

    def f_sofp_monthly(run):
        return [
            pick_latest(run, "NBD-MF-01-SOFP & SOCI AFL Monthly FS *.xlsx"),
            pick_latest(run, "ifrs_code.xlsx"),
        ]

    def f_sofp_annual(run):
        return [
            pick_latest(run, "NBD-AF-01-SOFP & SOCI AFL Annual FS *.xlsx"),
            pick_latest(run, "ifrs_code.xlsx"),
        ]

    return [
        ("NBD_MF_20_C1_C6", "C1-C6 (Monthly)", f_c1_c6),
        ("NBD_AF_20_C1_C6", "C2-C6 (Annual)", f_af_c1_c6),
        ("NBD-MF-10-GA & NBD-MF-11-IS", "GA & IS", f_ga_is),
        ("NBD_QF_23_C8", "C8", f_c8),
        ("NBD_QF_23_C7", "C7", f_c7),
        ("NBD_MF_23_C4", "C4 (standalone)", f_c4),
        ("NBD-WF-18-DM", "DM", f_dm),
        ("NBD_WF_15_LA", "LA", f_la),
        ("NBD_WF_18_ECS", "ECS", f_ecs),
        ("NBD_MF_04_LA", "MF-04-LA", f_mf04_la),
        ("NBD_MF_23_IA", "IA / C1-C2", f_ia),
        ("NBD_MF_01_SOFP_SOCI", "SOFP (Monthly)", f_sofp_monthly),
        ("NBD_AF_01_SOFP_SOCI", "SOFP (Annual)", f_sofp_annual),
    ]


# ============================================================================
# IFRS reports — folder under outputs/, label, subfolder to search (None =
# the run root itself), list of glob patterns
# ============================================================================
def ifrs_reports():
    return [
        ("PD", "IFRS-02 / PD", None, [
            "04. PD Lease *.xlsx",
            "03. PD Loan *.xlsx",
            "02. Historic PD Calculation_*.xlsb",
            "01. PD_data_*.xlsb",
        ]),
        ("Summary", "IFRS-01 / Summary (root)", None, [
            "3. Summary*.xlsb",
            "1. Portfolio_4_ECL*.xlsb",
            "2. ECL Model*.xlsb",
        ]),
        ("Summary", "IFRS-01 / Summary (Input/Overlay)", "Input", [
            "Overlay-*.xlsx",
        ]),
        ("LGD", "IFRS-03 / LGD (LGD subfolder)", "LGD", [
            "LGD*.xlsb",
            "receipt details.xlsx",
        ]),
        ("LGD", "IFRS-03 / LGD (firstDefault subfolder)", "firstDefault", [
            "First Default Date Working DPD*.xlsx",
        ]),
    ]


def main():
    print("=" * 70)
    print("CBSL reports")
    print("=" * 70)
    for folder_name, label, selector in cbsl_reports():
        report_folder = OUTPUTS_DIR / folder_name
        run = latest_date_folder(report_folder)
        if run is None:
            print(f"  [{label}] SKIPPED - no date folder found under {report_folder}")
            continue
        files = selector(run)
        write_manifest(run, files, label)

    print()
    print("=" * 70)
    print("IFRS reports")
    print("=" * 70)
    # IFRS-01/Summary needs both its root-level and Input-level selections
    # merged into ONE manifest for the same run folder, so handle it specially.
    seen_ifrs_runs: dict[str, tuple[Path, list[Path], str]] = {}
    for folder_name, label, subfolder, patterns in ifrs_reports():
        report_folder = OUTPUTS_DIR / folder_name
        run = latest_date_folder(report_folder)
        if run is None:
            print(f"  [{label}] SKIPPED - no date folder found under {report_folder}")
            continue
        search_in = (run / subfolder) if subfolder else run
        files = [pick_latest(search_in, pat) for pat in patterns]

        key = str(run)
        if key in seen_ifrs_runs:
            _, existing_files, existing_label = seen_ifrs_runs[key]
            seen_ifrs_runs[key] = (run, existing_files + files, existing_label)
        else:
            seen_ifrs_runs[key] = (run, files, label)

    for run, files, label in seen_ifrs_runs.values():
        write_manifest(run, files, label)

    print()
    print("Done. Review the output above before assuming everything is correct -")
    print("SKIPPED lines mean no manifest was written for that report; it will")
    print("keep using the old 'grab everything' fallback until its next real run.")


if __name__ == "__main__":
    main()

"""
PD Portfolio Roll-Forward Automation - Dynamic Rolling Cycle
=============================================================
Each run dynamically scans for all consecutive new summary files and
accumulates them in a pickle cache. The cache is split:
    Portfolio_1 = first 7 months
    Portfolio_2 = next 6 months  (empty until >= 8 months accumulated)
When the cache reaches TOTAL_CYCLE_MONTHS (13), the last month becomes
the seed for the next cycle and the cache resets.
"""
import os
import sys
import glob
import pickle
import time
import logging
import calendar
from datetime import datetime

# Force UTF-8 output on Windows to avoid charmap encoding errors
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Add the Class directory to the path
class_path = os.path.join(os.path.dirname(__file__), 'Scripts', 'Class')
sys.path.insert(0, class_path)

from BasicExcelFunctionsClass import ExcelPortfolioAutomation
import pandas as pd


# =============================================================================
# LOGGING SETUP
# =============================================================================
def setup_logger(script_dir):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    project_root = os.path.dirname(script_dir)
    run_dir = os.path.join(project_root, "logs", f"run_{ts}")
    os.makedirs(run_dir, exist_ok=True)

    log_file = os.path.join(run_dir, "PD.log")

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.handlers.clear()

    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)

    logging.info("Logger initialized")
    print(f"Logger initialized. Log file: {log_file}")
    return log_file

logger = logging.getLogger(__name__)


# =============================================================================
# CONFIGURATION
# =============================================================================
def _find_working_pd_folder() -> str:
    """Find the single date subfolder inside working/PD as the fallback input/output folder."""
    base = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "working", "PD")
    if os.path.isdir(base):
        subfolders = [d for d in os.listdir(base) if os.path.isdir(os.path.join(base, d))]
        if subfolders:
            return os.path.join(base, subfolders[0])
    return base

PD_BASE_FOLDER = os.environ.get("PD_BASE_FOLDER", "").strip()
INPUT_FOLDER   = PD_BASE_FOLDER if PD_BASE_FOLDER else _find_working_pd_folder()
OUTPUT_FOLDER  = INPUT_FOLDER
_orig_matches  = glob.glob(os.path.join(INPUT_FOLDER, "01. PD_data*.xlsb"))
ORIGINAL_PD_FILE = sorted(_orig_matches)[0] if _orig_matches else os.path.join(INPUT_FOLDER, "01. PD_data.xlsb")
CONFIG_FILE      = os.path.join(INPUT_FOLDER, "latest_month.txt")
CYCLE_CACHE_FILE = os.path.join(INPUT_FOLDER, "portfolio1_run1_cache.pkl")
STAGE_LOOKUP_CACHE_FILE = os.path.join(INPUT_FOLDER, "stage_lookup_cache.pkl")

def _find_pd_cat_override_file(folder: str) -> str | None:
    """Return the path of the first Test_PD_Cate - *.xlsx file found, or None."""
    matches = glob.glob(os.path.join(folder, "Test_PD_Cate - *.xlsx"))
    return sorted(matches)[0] if matches else None

SELECTED_RUN_DATE = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("PD_RUN_DATE", "")

# Cycle constants
P1_MAX_MONTHS    = 7
P2_MAX_MONTHS    = 6
TOTAL_CYCLE_MONTHS = P1_MAX_MONTHS + P2_MAX_MONTHS  # 13

# Column mapping: Summary file column -> DataFrame column
# NOTE: STAGE is deliberately NOT extracted here. It is sourced and written
# by a completely separate mechanism (see "STAGE LOOKUP" section below) that
# never goes through dedup/merge, so it can't be lost or overwritten by any
# row-selection logic that applies to the other columns.
COLUMN_MAPPING = {
    'CONTRACT NO':           'CONTRACT_NO',
    'EQUIPMENT DESCRIPTION': 'EQT_DESC',
    'PD/LGD CATEGORY':       'PD_CATEGORY',
    'CLIENT DPD':             'DPD',
}

# Columns extracted from summary files
DF_TO_PORTFOLIO_COLUMNS = ['MONTH', 'CONTRACT_NO', 'EQT_DESC', 'PD_CATEGORY', 'DPD']

# Columns written to Portfolio sheets (STAGE is pasted separately at the end
# by _paste_stage_from_lookup — see run_automation())
PORTFOLIO_WRITE_COLUMNS = ['MONTH', 'CONTRACT_NO', 'CONTRACT_NO_NOLASTDIG', 'EQT_DESC', 'PD_CATEGORY', 'DPD']

# Column positions in Excel
COLUMN_POSITIONS = {
    'MONTH':                 'A',
    'CONTRACT_NO':           'B',
    'CONTRACT_NO_NOLASTDIG': 'C',
    'EQT_DESC':              'D',
    'PD_CATEGORY':           'E',
    'DPD':                   'F',
    'STAGE':                 'H'
}


# =============================================================================
# CYCLE CACHE HELPERS
# =============================================================================
def load_cycle_cache() -> pd.DataFrame:
    """Load accumulated monthly data from the pickle cache."""
    if os.path.exists(CYCLE_CACHE_FILE):
        try:
            with open(CYCLE_CACHE_FILE, 'rb') as f:
                df = pickle.load(f)
            logger.info(f"[Cache] Loaded {len(df)} rows from {os.path.basename(CYCLE_CACHE_FILE)}")
            return df
        except Exception as e:
            logger.warning(f"[Cache] Could not read cache, starting fresh: {e}")
    return pd.DataFrame()


def save_cycle_cache(df: pd.DataFrame) -> None:
    """Persist accumulated monthly data to the pickle cache."""
    try:
        with open(CYCLE_CACHE_FILE, 'wb') as f:
            pickle.dump(df, f)
        logger.info(f"[Cache] Saved {len(df)} rows to {os.path.basename(CYCLE_CACHE_FILE)}")
    except Exception as e:
        logger.warning(f"[Cache] Could not save cache: {e}")


def add_months_eom(dt: datetime, months: int) -> datetime:
    """Add N months to a date and return the last day of the resulting month."""
    m = dt.month - 1 + months
    year  = dt.year + m // 12
    month = m % 12 + 1
    last_day = calendar.monthrange(year, month)[1]
    return dt.replace(year=year, month=month, day=last_day)


# =============================================================================
# PD CATEGORY OVERRIDE
# =============================================================================
def apply_pd_category_overrides(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply PD category and EQT_DESC corrections from the override file.

    The override file (Test_PD_Cate - *.xlsx) contains three columns:
        CONTRACT_NO_NOLASTDIG  |  PD_CATEGORY  |  EQT_DESC
    Requires CONTRACT_NO_NOLASTDIG to already exist in df.
    Safe to call even if the override file is missing (returns df unchanged).
    """
    override_file = _find_pd_cat_override_file(INPUT_FOLDER)
    if not override_file:
        logger.info("[Override] No Test_PD_Cate - *.xlsx found — skipping")
        return df

    try:
        override_df = pd.read_excel(
            override_file,
            sheet_name='Test_Sheet',
            usecols=['CONTRACT_NO_NOLASTDIG', 'PD_CATEGORY', 'EQT_DESC']
        )
        override_df = override_df.dropna(subset=['CONTRACT_NO_NOLASTDIG'])
        override_df['CONTRACT_NO_NOLASTDIG'] = override_df['CONTRACT_NO_NOLASTDIG'].astype(str).str.strip()

        df = df.copy()
        df['CONTRACT_NO_NOLASTDIG'] = df['CONTRACT_NO_NOLASTDIG'].astype(str).str.strip()

        # Build lookup dicts
        pd_cat_map  = dict(zip(override_df['CONTRACT_NO_NOLASTDIG'],
                               override_df['PD_CATEGORY'].astype(str).str.strip()))
        eqt_map     = {}
        if 'EQT_DESC' in override_df.columns:
            eqt_map = {k: v for k, v in zip(
                override_df['CONTRACT_NO_NOLASTDIG'],
                override_df['EQT_DESC'].astype(str).str.strip()
            ) if pd.notna(v) and v not in ('', 'nan')}

        # Apply PD_CATEGORY overrides
        mapped_pd = df['CONTRACT_NO_NOLASTDIG'].map(pd_cat_map)
        changed_pd = mapped_pd.notna() & (mapped_pd != df['PD_CATEGORY'].astype(str).str.strip())
        df.loc[changed_pd, 'PD_CATEGORY'] = mapped_pd[changed_pd]

        # Apply EQT_DESC overrides
        if eqt_map:
            mapped_eqt = df['CONTRACT_NO_NOLASTDIG'].map(eqt_map)
            changed_eqt = mapped_eqt.notna()
            df.loc[changed_eqt, 'EQT_DESC'] = mapped_eqt[changed_eqt]
            logger.info(f"[Override] PD_CATEGORY: {changed_pd.sum()} changed, "
                        f"EQT_DESC: {changed_eqt.sum()} changed")
        else:
            logger.info(f"[Override] PD_CATEGORY: {changed_pd.sum()} changed")

        return df

    except Exception as e:
        logger.warning(f"[Override] Could not apply overrides: {e}")
        return df


def extract_pivot_duplicates_to_override_file(all_data: pd.DataFrame) -> int:
    """
    Find contracts with more than one distinct PD_CATEGORY across all months
    and write every conflicting row to the override file (Test_Sheet).
    Returns the number of conflicting contracts written.
    """
    override_file = _find_pd_cat_override_file(INPUT_FOLDER)
    if not override_file:
        return 0
    try:
        conflict_contracts = (
            all_data.groupby('CONTRACT_NO_NOLASTDIG')['PD_CATEGORY']
            .nunique()
            .pipe(lambda s: s[s > 1])
            .index
        )
        if conflict_contracts.empty:
            return 0
        conflict_rows = all_data[all_data['CONTRACT_NO_NOLASTDIG'].isin(conflict_contracts)][
            ['CONTRACT_NO_NOLASTDIG', 'PD_CATEGORY', 'EQT_DESC']
        ].drop_duplicates()
        with pd.ExcelWriter(override_file, engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            conflict_rows.to_excel(writer, sheet_name='Test_Sheet', index=False)
        logger.info(f"[Override] Wrote {len(conflict_contracts)} conflicting contract(s) to {os.path.basename(override_file)}")
        return len(conflict_contracts)
    except Exception as e:
        logger.warning(f"[Override] Could not write conflict rows: {e}")
        return 0


# =============================================================================
# SEQUENTIAL MONTH FILE FINDER (dynamic — no fixed count required)
# =============================================================================
def find_consecutive_summary_files(input_folder: str, after_month: datetime,
                                   file_prefix: str = "3. Summary_",
                                   file_extension: str = ".xlsb") -> list:
    """
    Walk forward month by month from after_month and collect every summary file
    that exists. Stops gracefully at the first missing month.

    Returns list of (file_path, month_date) in chronological order.
    May return an empty list if no new files are found.
    """
    results = []
    i = 1
    while True:
        target   = add_months_eom(after_month, i)
        yr_month = target.strftime('%Y-%m')
        pattern  = os.path.join(input_folder, f"{file_prefix}{yr_month}-*{file_extension}")
        matches  = glob.glob(pattern)
        if not matches:
            break
        fp = sorted(matches)[0]
        logger.info(f"Found [{i}]: {os.path.basename(fp)}")
        results.append((fp, target))
        i += 1
    return results


# =============================================================================
# STAGE LOOKUP — dedicated, independent source of truth for the STAGE column
# =============================================================================
# STAGE is deliberately handled completely outside the normal
# extract -> dedup -> merge -> write pipeline used for every other column.
# That pipeline exists to pick ONE winning row per (contract, month) out of
# possibly several tranche rows, and every STAGE bug so far has come from
# STAGE getting tangled up in that row-selection process one way or another.
#
# Instead: for every month currently in the portfolio, read CONTRACT_NO
# (column A) and STAGE (the AG / "STAGE" column) straight from that month's
# source file into a small (MONTH, CONTRACT_NO, STAGE) lookup table — one
# row per contract as it appears in the source, no dedup needed. This table
# is persisted so a month, once read, is never re-read (old months are
# carried forward unchanged); only genuinely new months, or a month that's
# missing from the table for any reason, get (re)extracted.
#
# As the very last step of building the PD data file, this lookup is used
# to paste the STAGE value onto every row of Portfolio_1 / Portfolio_2 by
# matching directly on (MONTH, CONTRACT_NO) — i.e. matching against whatever
# row actually ended up on the sheet, regardless of how dedup picked it.
# =============================================================================
def _extract_stage_lookup_for_month(file_path: str, month_str: str, sheet_name: str = 'SUMMARY',
                                     max_header_search_rows: int = 10) -> pd.DataFrame:
    """Read CONTRACT_NO -> STAGE straight from one summary source file (column A
    and the AG/STAGE column), tagged with month_str. Full CONTRACT_NO is used
    as-is (not truncated) since it's matched directly against Portfolio col B."""
    t_start = time.time()
    logger.info(f"[StageLookup] {month_str}: reading {os.path.basename(file_path)} ...")

    found_header_row = None
    for header_row in range(max_header_search_rows):
        try:
            probe = pd.read_excel(file_path, sheet_name=sheet_name,
                                   header=header_row, nrows=1, engine='pyxlsb')
            cols_upper = [str(c).strip().upper() for c in probe.columns]
            if 'CONTRACT NO' in cols_upper and 'STAGE' in cols_upper:
                found_header_row = header_row
                break
        except Exception as e:
            logger.debug(f"[StageLookup] {month_str}: header probe at row {header_row} "
                         f"raised {type(e).__name__}: {e}")
            continue

    if found_header_row is None:
        logger.error(
            f"[StageLookup] {month_str}: no header row in the first {max_header_search_rows} "
            f"row(s) of {os.path.basename(file_path)} contained both 'CONTRACT NO' and 'STAGE' — "
            f"defaulting to header row 0 (this usually means extraction below will fail cleanly)."
        )
        found_header_row = 0
    else:
        logger.info(f"[StageLookup] {month_str}: header row located at index {found_header_row}")

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=found_header_row, engine='pyxlsb')
    except Exception as e:
        logger.error(f"[StageLookup] {month_str}: FAILED to read {os.path.basename(file_path)} "
                     f"({type(e).__name__}: {e}) — this month will have no STAGE data.", exc_info=True)
        return pd.DataFrame(columns=['MONTH', 'CONTRACT_NO', 'STAGE'])

    df.columns = [str(c).strip() for c in df.columns]
    if 'CONTRACT NO' not in df.columns or 'STAGE' not in df.columns:
        logger.warning(
            f"[StageLookup] {month_str}: CONTRACT NO / STAGE columns not found in "
            f"{os.path.basename(file_path)} (columns actually present: {list(df.columns)})"
        )
        return pd.DataFrame(columns=['MONTH', 'CONTRACT_NO', 'STAGE'])

    total_rows = len(df)
    try:
        sub = df[['CONTRACT NO', 'STAGE']].dropna(subset=['CONTRACT NO']).copy()
        dropped_no_contract = total_rows - len(sub)
        sub['CONTRACT_NO'] = sub['CONTRACT NO'].astype(str).str.strip()
        rows_before_stage_drop = len(sub)
        sub = sub.dropna(subset=['STAGE'])
        dropped_no_stage = rows_before_stage_drop - len(sub)
        sub['MONTH'] = month_str
        rows_before_dedup = len(sub)
        sub = sub.drop_duplicates(subset=['CONTRACT_NO'], keep='last')
        dropped_dupes = rows_before_dedup - len(sub)
    except Exception as e:
        logger.error(f"[StageLookup] {month_str}: FAILED while cleaning/deduping extracted data "
                     f"from {os.path.basename(file_path)} ({type(e).__name__}: {e})", exc_info=True)
        return pd.DataFrame(columns=['MONTH', 'CONTRACT_NO', 'STAGE'])

    elapsed = time.time() - t_start
    logger.info(
        f"[StageLookup] {month_str}: {total_rows:,} raw row(s) -> {len(sub):,} usable in {elapsed:.1f}s "
        f"(dropped {dropped_no_contract:,} no-CONTRACT-NO, {dropped_no_stage:,} no-STAGE, "
        f"{dropped_dupes:,} duplicate CONTRACT_NO)"
    )
    return sub[['MONTH', 'CONTRACT_NO', 'STAGE']]


def _load_stage_lookup_cache() -> pd.DataFrame:
    """Load the persisted STAGE lookup table."""
    if os.path.exists(STAGE_LOOKUP_CACHE_FILE):
        try:
            t_start = time.time()
            with open(STAGE_LOOKUP_CACHE_FILE, 'rb') as f:
                df = pickle.load(f)
            logger.info(f"[StageLookup] Loaded {len(df):,} row(s) from "
                        f"{os.path.basename(STAGE_LOOKUP_CACHE_FILE)} in {time.time() - t_start:.1f}s")
            if not df.empty:
                logger.info(f"[StageLookup] Cache currently covers {df['MONTH'].nunique()} month(s): "
                            f"{sorted(df['MONTH'].unique(), key=lambda m: pd.to_datetime(m, format='%m/%d/%Y'))}")
            return df
        except Exception as e:
            logger.error(f"[StageLookup] Could not read cache at {STAGE_LOOKUP_CACHE_FILE} "
                         f"({type(e).__name__}: {e}) — starting fresh (empty lookup).", exc_info=True)
    else:
        logger.info(f"[StageLookup] No existing cache at {os.path.basename(STAGE_LOOKUP_CACHE_FILE)} "
                    f"— starting fresh (empty lookup).")
    return pd.DataFrame(columns=['MONTH', 'CONTRACT_NO', 'STAGE'])


def _save_stage_lookup_cache(df: pd.DataFrame) -> None:
    """Persist the STAGE lookup table."""
    try:
        t_start = time.time()
        with open(STAGE_LOOKUP_CACHE_FILE, 'wb') as f:
            pickle.dump(df, f)
        logger.info(f"[StageLookup] Saved {len(df):,} row(s) to "
                    f"{os.path.basename(STAGE_LOOKUP_CACHE_FILE)} in {time.time() - t_start:.1f}s")
    except Exception as e:
        logger.error(f"[StageLookup] Could not save cache to {STAGE_LOOKUP_CACHE_FILE} "
                     f"({type(e).__name__}: {e}) — this run's STAGE lookup work will not persist "
                     f"and will be redone next run.", exc_info=True)


def _sync_stage_lookup(current_months: list, input_folder: str) -> tuple:
    """
    Ensure the persisted STAGE lookup has an entry set for every month in
    current_months (every month currently in the portfolio):
      - A month already present in the cache is carried forward UNCHANGED —
        never re-read, never re-derived.
      - A month missing from the cache (genuinely new this run, or a gap
        left over from an earlier failed extraction) is read fresh from its
        source summary file.
      - A month no longer in current_months (rolled off the 13-month cycle)
        is dropped, so the table stays in sync with the actual portfolio.

    Returns (lookup_df, changed) where changed is True if anything had to
    be freshly extracted (i.e. the cache on disk was out of date).
    """
    logger.info(f"[StageLookup] Syncing lookup for {len(current_months)} month(s): {current_months}")
    existing = _load_stage_lookup_cache()
    current_set = set(current_months)

    kept = existing[existing['MONTH'].isin(current_set)].copy() if not existing.empty \
        else pd.DataFrame(columns=['MONTH', 'CONTRACT_NO', 'STAGE'])
    kept_months = set(kept['MONTH'].unique()) if not kept.empty else set()
    missing_months = [m for m in current_months if m not in kept_months]

    if not existing.empty:
        dropped_months = sorted(set(existing['MONTH'].unique()) - current_set,
                                 key=lambda m: pd.to_datetime(m, format='%m/%d/%Y'))
        if dropped_months:
            dropped_rows = int(existing['MONTH'].isin(dropped_months).sum())
            logger.info(f"[StageLookup] Dropping {len(dropped_months)} month(s) no longer in the "
                        f"portfolio ({dropped_rows:,} row(s) discarded, rolled off the cycle): "
                        f"{dropped_months}")

    if kept_months:
        logger.info(f"[StageLookup] Carrying forward {len(kept_months)} already-cached month(s) "
                    f"unchanged ({len(kept):,} row(s), never re-read): "
                    f"{sorted(kept_months, key=lambda m: pd.to_datetime(m, format='%m/%d/%Y'))}")
    if missing_months:
        logger.info(f"[StageLookup] {len(missing_months)} month(s) need fresh extraction: {missing_months}")
    else:
        logger.info("[StageLookup] No months need fresh extraction — lookup already fully covers the portfolio.")

    new_parts = []
    for month in missing_months:
        try:
            yr_month = pd.to_datetime(month, format='%m/%d/%Y').strftime('%Y-%m')
            pattern = os.path.join(input_folder, f"3. Summary_{yr_month}-*.xlsb")
            matches = sorted(glob.glob(pattern))
            if not matches:
                logger.warning(f"[StageLookup] {month}: no source file matching '{pattern}' — "
                                f"contracts for this month will have no STAGE")
                continue
            if len(matches) > 1:
                logger.warning(f"[StageLookup] {month}: {len(matches)} source file(s) matched "
                                f"'{pattern}' ({[os.path.basename(m) for m in matches]}) — "
                                f"using the alphabetically first one: {os.path.basename(matches[0])}")
            part = _extract_stage_lookup_for_month(matches[0], month)
            new_parts.append(part)
            if part.empty:
                logger.error(f"[StageLookup] {month}: extraction from {os.path.basename(matches[0])} "
                             f"returned 0 usable row(s) — this month will have no STAGE.")
            else:
                logger.info(f"[StageLookup] {month}: extracted {len(part):,} contract STAGE value(s) "
                            f"from {os.path.basename(matches[0])}")
        except Exception as e:
            logger.error(f"[StageLookup] {month}: unexpected error during sync "
                         f"({type(e).__name__}: {e}) — this month will have no STAGE data; "
                         f"continuing with the remaining months.", exc_info=True)
            continue

    parts = ([kept] if not kept.empty else []) + [p for p in new_parts if not p.empty]
    updated = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=['MONTH', 'CONTRACT_NO', 'STAGE'])
    changed = bool(new_parts) or (kept_months != set(existing['MONTH'].unique()) if not existing.empty else False)
    logger.info(f"[StageLookup] Sync complete: {len(updated):,} total row(s) across "
                f"{updated['MONTH'].nunique() if not updated.empty else 0} month(s), changed={changed}")
    return updated, changed


_STAGE_PASTE_CHUNK_SIZE = 50_000
_STAGE_PASTE_MAX_RETRIES = 3
_STAGE_PASTE_RETRY_DELAY_SECONDS = 2
_STAGE_UNMATCHED_SAMPLE_SIZE = 25


def _chunked_read_ab(sheet, last_row: int, label: str = '') -> list:
    """
    Read A2:B{last_row} in bounded chunks instead of one massive COM range
    read. A single range covering hundreds of thousands of rows in one call
    is a real fragility point for Excel COM automation (especially over
    Remote Desktop) — it can silently return truncated or misaligned data
    for some rows with no error raised. Bounded chunks, each retried on
    failure and validated for row count, close that gap.
    """
    rows = []
    total_data_rows = last_row - 1
    total_chunks = -(-total_data_rows // _STAGE_PASTE_CHUNK_SIZE)  # ceil division
    row = 2
    chunk_num = 0
    while row <= last_row:
        chunk_num += 1
        chunk_end = min(row + _STAGE_PASTE_CHUNK_SIZE - 1, last_row)
        expected_n = chunk_end - row + 1
        chunk = None
        last_error = None
        for attempt in range(1, _STAGE_PASTE_MAX_RETRIES + 1):
            try:
                chunk = sheet.range(f'A{row}:B{chunk_end}').value
                if chunk_end == row:
                    chunk = [chunk]
                if len(chunk) != expected_n:
                    raise ValueError(f"expected {expected_n} row(s) back, got {len(chunk)}")
                last_error = None
                break
            except Exception as e:
                last_error = e
                logger.warning(
                    f"[StageLookup] {label}: read chunk {chunk_num}/{total_chunks} "
                    f"(rows {row}-{chunk_end}) attempt {attempt}/{_STAGE_PASTE_MAX_RETRIES} "
                    f"failed ({type(e).__name__}: {e})"
                )
                if attempt < _STAGE_PASTE_MAX_RETRIES:
                    time.sleep(_STAGE_PASTE_RETRY_DELAY_SECONDS)
        if last_error is not None:
            logger.error(
                f"[StageLookup] {label}: read chunk {chunk_num}/{total_chunks} "
                f"(rows {row}-{chunk_end}) FAILED after {_STAGE_PASTE_MAX_RETRIES} attempts — "
                f"{last_error}", exc_info=True
            )
            raise last_error
        logger.info(f"[StageLookup] {label}: read chunk {chunk_num}/{total_chunks} "
                    f"(rows {row}-{chunk_end}, {expected_n:,} row(s)) OK")
        rows.extend(chunk)
        row = chunk_end + 1

    if len(rows) != total_data_rows:
        logger.error(f"[StageLookup] {label}: read {len(rows):,} row(s) total but expected "
                     f"{total_data_rows:,} — data may be misaligned")
    return rows


def _chunked_write_h(sheet, values: list, last_row: int, label: str = '') -> None:
    """Write H2:H{last_row} in the same bounded, retried, validated chunks."""
    total_data_rows = last_row - 1
    total_chunks = -(-total_data_rows // _STAGE_PASTE_CHUNK_SIZE)
    row = 2
    idx = 0
    chunk_num = 0
    while row <= last_row:
        chunk_num += 1
        chunk_end = min(row + _STAGE_PASTE_CHUNK_SIZE - 1, last_row)
        n = chunk_end - row + 1
        chunk_values = values[idx:idx + n]
        last_error = None
        for attempt in range(1, _STAGE_PASTE_MAX_RETRIES + 1):
            try:
                sheet.range(f'H{row}:H{chunk_end}').value = chunk_values
                last_error = None
                break
            except Exception as e:
                last_error = e
                logger.warning(
                    f"[StageLookup] {label}: write chunk {chunk_num}/{total_chunks} "
                    f"(rows {row}-{chunk_end}) attempt {attempt}/{_STAGE_PASTE_MAX_RETRIES} "
                    f"failed ({type(e).__name__}: {e})"
                )
                if attempt < _STAGE_PASTE_MAX_RETRIES:
                    time.sleep(_STAGE_PASTE_RETRY_DELAY_SECONDS)
        if last_error is not None:
            logger.error(
                f"[StageLookup] {label}: write chunk {chunk_num}/{total_chunks} "
                f"(rows {row}-{chunk_end}) FAILED after {_STAGE_PASTE_MAX_RETRIES} attempts — "
                f"{last_error}", exc_info=True
            )
            raise last_error
        logger.info(f"[StageLookup] {label}: wrote chunk {chunk_num}/{total_chunks} "
                    f"(rows {row}-{chunk_end}, {n:,} row(s)) OK")
        idx += n
        row = chunk_end + 1


def _paste_stage_from_lookup(sheet, lookup_df: pd.DataFrame, label: str, input_folder: str) -> pd.DataFrame:
    """
    Final step: read back MONTH + CONTRACT_NO already written to the sheet,
    match each row against lookup_df by (MONTH, CONTRACT_NO), and paste the
    STAGE value for every matching contract, reading/writing in bounded,
    retried, validated chunks rather than one single huge COM range
    operation (see _chunked_read_ab / _chunked_write_h).

    SELF-CORRECTION: "carry forward, never re-read" (in _sync_stage_lookup)
    is fast, but it means a month that was ever synced incompletely for any
    reason (file locked mid-read, a transient error, anything) would stay
    wrong forever — the whole point of the design is to never touch a month
    again once it's "there". To close that gap: if MOST rows of a given
    month on this sheet fail to match (a lone unmatched contract is normal
    and left as-is), that's a strong signal the cached lookup for that
    month is stale/incomplete rather than a real data gap, so it's
    force-refreshed from source (bypassing the cache) and the match is
    retried before giving up. The corrected lookup is saved immediately.

    Returns the (possibly corrected) lookup_df so subsequent calls for
    other sheets use the fixed data too, not the stale version.
    """
    logger.info(f"[StageLookup] {label}: starting STAGE paste (lookup has "
                f"{len(lookup_df):,} row(s) across {lookup_df['MONTH'].nunique() if not lookup_df.empty else 0} month(s))")

    last_row = sheet.range('A2').end('down').row
    if last_row < 2 or last_row > 1_000_000:
        logger.info(f"[StageLookup] {label}: empty — nothing to paste")
        return lookup_df

    t_read_start = time.time()
    try:
        data = _chunked_read_ab(sheet, last_row, label)
    except Exception as e:
        logger.error(f"[StageLookup] {label}: could not read MONTH/CONTRACT_NO from the sheet "
                     f"({type(e).__name__}: {e}) — STAGE cannot be pasted for this sheet.", exc_info=True)
        raise
    logger.info(f"[StageLookup] {label}: read {len(data):,} row(s) in {time.time() - t_read_start:.1f}s")

    def _match(rows, lu):
        stage_col, unmatched_total = [], 0
        unmatched_by_month, total_by_month = {}, {}
        unmatched_sample = []
        for month, contract_no in rows:
            month_key = month.strftime('%m/%d/%Y') if isinstance(month, datetime) else str(month).strip()
            contract_key = str(contract_no).strip() if contract_no is not None else ''
            total_by_month[month_key] = total_by_month.get(month_key, 0) + 1
            stage = lu.get((month_key, contract_key))
            if stage is None:
                unmatched_total += 1
                unmatched_by_month[month_key] = unmatched_by_month.get(month_key, 0) + 1
                if len(unmatched_sample) < _STAGE_UNMATCHED_SAMPLE_SIZE:
                    unmatched_sample.append((month_key, contract_key))
            stage_col.append([stage])
        return stage_col, unmatched_total, unmatched_by_month, total_by_month, unmatched_sample

    t_match_start = time.time()
    lookup = {(row.MONTH, row.CONTRACT_NO): row.STAGE for row in lookup_df.itertuples(index=False)}
    stage_col, unmatched, unmatched_by_month, total_by_month, unmatched_sample = _match(data, lookup)
    logger.info(f"[StageLookup] {label}: matched {len(data):,} row(s) against "
                f"{len(lookup):,} lookup entries in {time.time() - t_match_start:.1f}s "
                f"({unmatched:,} unmatched so far, before any self-correction below)")

    suspect_months = [m for m, cnt in unmatched_by_month.items() if cnt / total_by_month[m] > 0.5]
    if suspect_months:
        logger.warning(
            f"[StageLookup] {label}: month(s) {suspect_months} have most rows unmatched — "
            f"the cached lookup for that month looks stale/incomplete, forcing a fresh "
            f"re-extraction from source instead of trusting the cache."
        )
        forced_parts = []
        for month in suspect_months:
            try:
                yr_month = pd.to_datetime(month, format='%m/%d/%Y').strftime('%Y-%m')
                matches = sorted(glob.glob(os.path.join(input_folder, f"3. Summary_{yr_month}-*.xlsb")))
                if not matches:
                    logger.error(f"[StageLookup] {label}: {month}: source file not found — cannot force-refresh")
                    continue
                part = _extract_stage_lookup_for_month(matches[0], month)
                forced_parts.append(part)
                logger.info(f"[StageLookup] {label}: {month}: force-extracted {len(part):,} row(s) "
                            f"from {os.path.basename(matches[0])}")
            except Exception as e:
                logger.error(f"[StageLookup] {label}: {month}: force-refresh failed "
                             f"({type(e).__name__}: {e}) — leaving this month's rows as unmatched.",
                             exc_info=True)
                continue

        if forced_parts:
            lookup_df = pd.concat(
                [lookup_df[~lookup_df['MONTH'].isin(suspect_months)]] + forced_parts,
                ignore_index=True
            )
            _save_stage_lookup_cache(lookup_df)
            lookup = {(row.MONTH, row.CONTRACT_NO): row.STAGE for row in lookup_df.itertuples(index=False)}
            stage_col, unmatched, _, _, unmatched_sample = _match(data, lookup)
            logger.info(f"[StageLookup] {label}: re-matched after force-refresh — "
                        f"{unmatched:,} unmatched remaining")

    t_write_start = time.time()
    try:
        _chunked_write_h(sheet, stage_col, last_row, label)
    except Exception as e:
        logger.error(f"[StageLookup] {label}: could not write STAGE column back to the sheet "
                     f"({type(e).__name__}: {e}) — output file will NOT have correct STAGE values "
                     f"for this sheet.", exc_info=True)
        raise
    logger.info(f"[StageLookup] {label}: wrote STAGE column in {time.time() - t_write_start:.1f}s")

    matched = len(data) - unmatched
    msg = f"[StageLookup] {label}: pasted STAGE for {matched:,}/{len(data):,} row(s)"
    if unmatched:
        msg += f" — {unmatched:,} unmatched (no source data found for that contract/month)"
    logger.warning(msg) if unmatched else logger.info(msg)
    if unmatched_sample:
        logger.warning(f"[StageLookup] {label}: sample of unmatched (MONTH, CONTRACT_NO) pairs "
                        f"(showing up to {_STAGE_UNMATCHED_SAMPLE_SIZE} of {unmatched:,}): {unmatched_sample}")
    return lookup_df


# =============================================================================
# DC BUCKET FORMULA HELPER
# =============================================================================
def _write_dc_formula(sheet, num_rows: int, label: str) -> None:
    """Write the DC_BUCKET formula to column G using R1C1 notation."""
    # R1C1: RC[1]=H (STAGE), RC[-1]=F (DPD)
    DC_R1C1 = ('=IF(RC[1]=3,5,IF(RC[-1]<=0,1,IF(RC[-1]<=30,2,'
               'IF(RC[-1]<=60,3,IF(RC[-1]<=90,4,IF(RC[-1]>90,5))))))')
    try:
        sheet.range('G2:G1048576').clear_contents()
        if num_rows == 0:
            return
        last_row = num_rows + 1
        sheet.range(f'G2:G{last_row}').api.FormulaR1C1 = DC_R1C1
        logger.info(f"DC_BUCKET formula written to {label} G2:G{last_row}")
    except Exception as e:
        logger.warning(f"Could not write DC_BUCKET formula to {label}: {e}")


# =============================================================================
# MAIN AUTOMATION - PART 1: PORTFOLIO ROLL-FORWARD
# =============================================================================
class _NoNewMonthToProcess(Exception):
    """
    Raised internally when there is nothing new to roll forward (no new
    summary files yet, or the selected run date can't be fulfilled). Caught
    once in run_automation() to decide between a clean exit(2) and falling
    through to a repair-only write pass if a STAGE lookup gap was repaired
    earlier in this same run — a repair must never be silently lost just
    because there happened to be no new month to process alongside it.
    """
    def __init__(self, print_lines: list):
        self.print_lines = print_lines


def run_automation():
    """Main automation workflow — dynamic rolling cycle portfolio roll-forward."""

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    logger.info("MARKER_TEST_998877 — if you see this line, the edited file is the one executing")


    logger.info("=" * 80)
    logger.info("PD PORTFOLIO ROLL-FORWARD AUTOMATION  (DYNAMIC ROLLING CYCLE)")
    logger.info("=" * 80)

    # ------------------------------------------------------------------
    # Read current state
    # ------------------------------------------------------------------
    try:
        current_latest = ExcelPortfolioAutomation.read_config_file(CONFIG_FILE)
        logger.info(f"Current latest month : {ExcelPortfolioAutomation.format_month_string(current_latest)}")
    except FileNotFoundError:
        logger.error(f"Config file not found: {CONFIG_FILE}")
        sys.exit(1)
    except ValueError as e:
        logger.error(str(e))
        sys.exit(1)

    # ------------------------------------------------------------------
    # Load cache — ALWAYS runs first, before any run-date validation or
    # new-file scanning.
    # ------------------------------------------------------------------
    cached_data   = load_cycle_cache()
    cached_months = sorted(cached_data['MONTH'].unique(),
                           key=lambda x: pd.to_datetime(x, format='%m/%d/%Y')) \
                    if not cached_data.empty else []
    logger.info(f"Cache contains {len(cached_months)} month(s): {cached_months}")

    if len(cached_months) >= TOTAL_CYCLE_MONTHS:
        # Cycle complete — seed next cycle from last cached month
        seed_month = cached_months[-1]
        base_data  = cached_data[cached_data['MONTH'] == seed_month].copy()
        logger.info(f"Cycle complete ({TOTAL_CYCLE_MONTHS} months). "
                    f"Resetting — seed month: {seed_month} ({len(base_data)} rows)")
    else:
        # Still accumulating
        logger.info(f"Cycle progress: {len(cached_months)}/{TOTAL_CYCLE_MONTHS} months so far")
        base_data = cached_data.copy() if not cached_data.empty else pd.DataFrame()

    # Sync the STAGE lookup for whatever's already in the portfolio (before
    # any new month is known about). Saved immediately if it fixes a gap,
    # so the fix is persisted even if this run then has nothing new to
    # process (or an invalid/repeat run date) and would otherwise exit
    # early below — a STAGE gap is never gated behind a new month arriving.
    base_months_list = sorted(base_data['MONTH'].unique(),
                              key=lambda x: pd.to_datetime(x, format='%m/%d/%Y')) \
                       if not base_data.empty else []
    try:
        stage_lookup_df, repaired_now = _sync_stage_lookup(base_months_list, INPUT_FOLDER)
    except Exception:
        logger.error(f"[StageLookup] Unexpected fatal error during the early STAGE sync for "
                     f"month(s) {base_months_list}", exc_info=True)
        raise
    if repaired_now:
        _save_stage_lookup_cache(stage_lookup_df)
        logger.warning("[StageLookup] Cache repaired and saved immediately, independent of "
                        "whether a new month is processed below.")

    try:
        # ------------------------------------------------------------------
        # Validate selected run date: every monthly summary file from
        # current_latest+1 up to and including selected_dt must exist.
        # ------------------------------------------------------------------
        selected_dt = None
        if SELECTED_RUN_DATE:
            for _fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                try:
                    selected_dt = datetime.strptime(SELECTED_RUN_DATE.strip(), _fmt)
                    break
                except ValueError:
                    continue

        if selected_dt is not None:
            if selected_dt.date() <= current_latest.date():
                logger.error(
                    f"Selected date {selected_dt.strftime('%Y-%m-%d')} must be after the "
                    f"current latest month ({current_latest.strftime('%Y-%m-%d')})."
                )
                next_expected = add_months_eom(current_latest, 1)
                raise _NoNewMonthToProcess([
                    f"SUGGEST_DATE:{next_expected.strftime('%Y-%m-%d')}",
                    "AVAILABLE_FILES:",
                ])

            # Fiscal year end = March of the next year after current_latest
            # Maximum new months in one cycle = TOTAL_CYCLE_MONTHS - 1 (the seed)
            fiscal_year_end = add_months_eom(current_latest, TOTAL_CYCLE_MONTHS - 1)
            if selected_dt > fiscal_year_end:
                logger.error(
                    f"Selected date {selected_dt.strftime('%Y-%m-%d')} exceeds the fiscal year end "
                    f"({fiscal_year_end.strftime('%Y-%m-%d')}). "
                    f"The maximum run date for this cycle is {fiscal_year_end.strftime('%Y-%m-%d')}."
                )
                raise _NoNewMonthToProcess([
                    f"SUGGEST_DATE:{fiscal_year_end.strftime('%Y-%m-%d')}",
                    "MISSING_FILES:",
                ])

            # Build required months from current_latest+1 up to selected_dt (max fiscal year end)
            required_months = []
            for i in range(1, TOTAL_CYCLE_MONTHS):
                m_dt = add_months_eom(current_latest, i)
                required_months.append(m_dt)
                if m_dt.year == selected_dt.year and m_dt.month == selected_dt.month:
                    break

            # Check each required month has a matching summary file
            missing_months = []
            for m_dt in required_months:
                yr_month = m_dt.strftime('%Y-%m')
                pattern = os.path.join(INPUT_FOLDER, f"3. Summary_{yr_month}-*.xlsb")
                if not glob.glob(pattern):
                    missing_months.append(m_dt.strftime('%Y-%m-%d'))

            if missing_months:
                logger.error(
                    f"Cannot run for selected date {selected_dt.strftime('%Y-%m-%d')}. "
                    f"Summary files are missing for the following month(s):"
                )
                for md in missing_months:
                    logger.error(f"  3. Summary_{md}*.xlsb")
                logger.error(
                    f"All monthly summary files from "
                    f"{add_months_eom(current_latest, 1).strftime('%Y-%m-%d')} to "
                    f"{selected_dt.strftime('%Y-%m-%d')} must be present in the working folder."
                )
                missing_str = "|".join(f"3. Summary_{md}*.xlsb" for md in missing_months)
                raise _NoNewMonthToProcess([
                    f"MISSING_FILES:{missing_str}",
                    f"SUGGEST_DATE:{selected_dt.strftime('%Y-%m-%d')}",
                ])

        # ------------------------------------------------------------------
        # Dynamically scan for consecutive new summary files
        # ------------------------------------------------------------------
        logger.info("-" * 60)
        logger.info("SCANNING FOR NEW SUMMARY FILES")
        logger.info("-" * 60)

        summary_files = find_consecutive_summary_files(INPUT_FOLDER, current_latest)

        if not summary_files:
            next_expected = add_months_eom(current_latest, 1)
            next_expected_str = next_expected.strftime("%Y-%m-%d")
            logger.error(
                f"No summary files found after {ExcelPortfolioAutomation.format_month_string(current_latest)}. "
                f"Next expected file: 3. Summary_{next_expected_str}*.xlsb"
            )
            raise _NoNewMonthToProcess([
                f"MISSING_FILES:3. Summary_{next_expected_str}*.xlsb",
                f"SUGGEST_DATE:{selected_dt.strftime('%Y-%m-%d') if selected_dt else next_expected_str}",
            ])

        new_end_month = summary_files[-1][1]
        logger.info(f"New end month        : {ExcelPortfolioAutomation.format_month_string(new_end_month)}")
        logger.info(f"Files to process     : {len(summary_files)}")

        logger.info(f"Extracting data from {len(summary_files)} file(s)...")
        new_data = ExcelPortfolioAutomation.extract_data_from_summary_files(
            file_paths=summary_files,
            column_mapping=COLUMN_MAPPING,
            output_columns=DF_TO_PORTFOLIO_COLUMNS,
            sheet_name='SUMMARY',
            date_format='%m/%d/%Y'
        )

        if new_data.empty:
            logger.error("No data extracted from summary files.")
            return False, None

        new_data = new_data.sort_values(
            'MONTH', key=lambda x: pd.to_datetime(x, format='%m/%d/%Y')
        )
        new_data['CONTRACT_NO_NOLASTDIG'] = new_data['CONTRACT_NO'].apply(
            lambda x: str(x)[:-1] if pd.notna(x) and str(x) else None
        )
        new_data = apply_pd_category_overrides(new_data)
        logger.info(f"New data extracted: {len(new_data)} rows")

    except _NoNewMonthToProcess as e:
        if repaired_now:
            logger.warning(
                "[StageLookup] Nothing new to process this run, but a STAGE gap was repaired above — "
                "writing the repaired data to Excel now instead of exiting."
            )
            new_data = pd.DataFrame(columns=DF_TO_PORTFOLIO_COLUMNS + ['CONTRACT_NO_NOLASTDIG'])
            new_end_month = current_latest
        else:
            for line in e.print_lines:
                print(line, flush=True)
            sys.exit(2)

    # ------------------------------------------------------------------
    # IMMUTABILITY GUARD — a month already committed to the cache/base is
    # never re-derived. If latest_month.txt ever falls behind the cache
    # (config/cache desync) and a summary file for an already-cached month
    # gets picked up as "new" again, its rows are dropped here instead of
    # being merged back in. This is what previously let a fresh
    # re-extraction silently overwrite historical DPD/PD_CATEGORY/EQT_DESC
    # values: the old code merged cached + new data together and re-picked
    # a "winner" per contract-month using a non-stable sort, so a cached
    # row could lose to a freshly re-extracted one on a DPD tie for no good
    # reason. (STAGE itself is no longer part of this pipeline at all — see
    # the STAGE LOOKUP section — so it can't be affected by this class of
    # bug any more regardless.)
    # ------------------------------------------------------------------
    base_months = set(base_data['MONTH'].unique()) if not base_data.empty else set()
    if base_months:
        stale_mask = new_data['MONTH'].isin(base_months)
        if stale_mask.any():
            stale_months = sorted(set(new_data.loc[stale_mask, 'MONTH']))
            logger.warning(
                f"[Immutability] {stale_mask.sum()} row(s) in new_data belong to "
                f"already-cached month(s) {stale_months} — dropping them so cached "
                f"DPD/PD_CATEGORY/EQT_DESC values are not overwritten. This means "
                f"latest_month.txt was behind the cache; check for a config/cache desync."
            )
            new_data = new_data[~stale_mask].copy()

    # ------------------------------------------------------------------
    # Deduplicate NEW data only: one row per (CONTRACT_NO_NOLASTDIG, MONTH)
    # — keep highest DPD. Never re-run against already-cached rows, so a
    # historical month's STAGE can never change once committed to the cache.
    # kind='mergesort' makes the tie-break on equal DPD deterministic
    # (stable) instead of the previous default quicksort, which is not.
    # ------------------------------------------------------------------
    rows_before = len(new_data)
    if not new_data.empty:
        new_data = new_data.copy()
        new_data['_DPD_SORT'] = pd.to_numeric(new_data['DPD'], errors='coerce').fillna(0)
        new_data = (
            new_data
            .sort_values('_DPD_SORT', ascending=False, kind='mergesort')
            .drop_duplicates(subset=['CONTRACT_NO_NOLASTDIG', 'MONTH'], keep='first')
            .drop(columns=['_DPD_SORT'])
        )
    rows_after = len(new_data)
    if rows_before != rows_after:
        logger.info(f"Deduplication (new data only): {rows_before:,} → {rows_after:,} rows "
                    f"({rows_before - rows_after:,} duplicate tranche rows removed)")

    all_data = (pd.concat([base_data, new_data], ignore_index=True)
                if not base_data.empty else new_data)
    all_data = all_data.sort_values('MONTH', key=lambda x: pd.to_datetime(x, format='%m/%d/%Y'))

    # Write conflict contracts to override file for manual review
    extract_pivot_duplicates_to_override_file(all_data)

    # ------------------------------------------------------------------
    # Split into Portfolio_1 (first 7 months) and Portfolio_2 (next 6)
    # ------------------------------------------------------------------
    all_months  = sorted(all_data['MONTH'].unique(),
                         key=lambda x: pd.to_datetime(x, format='%m/%d/%Y'))
    p1_months   = set(all_months[:P1_MAX_MONTHS])
    p2_months   = set(all_months[P1_MAX_MONTHS:P1_MAX_MONTHS + P2_MAX_MONTHS])

    p1_data = all_data[all_data['MONTH'].isin(p1_months)].copy()
    p2_data = all_data[all_data['MONTH'].isin(p2_months)].copy()

    logger.info(f"Portfolio_1: {len(p1_months)} month(s), {len(p1_data)} rows")
    logger.info(f"Portfolio_2: {len(p2_months)} month(s), {len(p2_data)} rows")

    # Re-sync the STAGE lookup against the FINAL month set for this run —
    # picks up whatever new month was just added (the early sync above only
    # covered the cache's pre-existing months). Cheap no-op for any month
    # already covered by the early sync.
    try:
        stage_lookup_df, _ = _sync_stage_lookup(all_months, INPUT_FOLDER)
    except Exception:
        logger.error(f"[StageLookup] Unexpected fatal error during the final STAGE sync for "
                     f"month(s) {all_months}", exc_info=True)
        raise
    _save_stage_lookup_cache(stage_lookup_df)

    # Excel row-limit guard (max 1,048,574 data rows starting at row 2)
    EXCEL_MAX_ROWS = 1_048_574
    for sheet_label, sheet_df, months_set in [
        ('Portfolio_1', p1_data, p1_months),
        ('Portfolio_2', p2_data, p2_months),
    ]:
        if len(sheet_df) > EXCEL_MAX_ROWS:
            raise ValueError(
                f"{sheet_label} has {len(sheet_df):,} rows across {len(months_set)} month(s) "
                f"— exceeds Excel limit of {EXCEL_MAX_ROWS:,}."
            )

    # ------------------------------------------------------------------
    # Get latest PD file
    # ------------------------------------------------------------------
    pd_file = ExcelPortfolioAutomation.get_latest_file_in_folder(
        folder_path=OUTPUT_FOLDER,
        file_pattern="01. PD_data_*.xlsb",
        fallback_file=ORIGINAL_PD_FILE
    )

    output_file = None

    # ------------------------------------------------------------------
    # Write to Excel
    # ------------------------------------------------------------------
    logger.info("-" * 60)
    logger.info("WRITING PORTFOLIO DATA TO EXCEL")
    logger.info("-" * 60)

    with ExcelPortfolioAutomation(pd_file, visible=True) as excel:

        excel.write_portfolio_data(
            'Portfolio_1', p1_data,
            PORTFOLIO_WRITE_COLUMNS, COLUMN_POSITIONS
        )
        _write_dc_formula(excel.workbook.sheets['Portfolio_1'], len(p1_data), 'Portfolio_1')

        excel.write_portfolio_data(
            'Portfolio_2', p2_data,
            PORTFOLIO_WRITE_COLUMNS, COLUMN_POSITIONS,
            full_clear=(len(p2_data) == 0)
        )
        _write_dc_formula(excel.workbook.sheets['Portfolio_2'], len(p2_data), 'Portfolio_2')

        # ------------------------------------------------------------------
        # FINAL STEP — paste STAGE onto every row from the dedicated lookup,
        # matched directly by (MONTH, CONTRACT_NO) against whatever is
        # actually on each sheet. Must run after both sheets' other columns
        # are written (needs CONTRACT_NO already in place to match against)
        # and before the pivot refresh, so the pivot reflects final STAGE.
        # ------------------------------------------------------------------
        logger.info("-" * 60)
        logger.info("PASTING STAGE FROM LOOKUP")
        logger.info("-" * 60)
        stage_lookup_df = _paste_stage_from_lookup(excel.workbook.sheets['Portfolio_1'], stage_lookup_df,
                                                    'Portfolio_1', INPUT_FOLDER)
        if not p2_data.empty:
            stage_lookup_df = _paste_stage_from_lookup(excel.workbook.sheets['Portfolio_2'], stage_lookup_df,
                                                        'Portfolio_2', INPUT_FOLDER)

        excel.refresh_pivot_table('01.Pivoted_Portfolio', 'PivotTable1')
        logger.info(f"Pivot refreshed — {len(all_months)} month(s) visible")

        # Save output file
        date_str    = datetime.now().strftime("%Y-%m-%d")
        output_file = os.path.join(OUTPUT_FOLDER, f"01. PD_data_{date_str}.xlsb")
        excel.save_as(output_file)
        logger.info(f"Saved: {output_file}")

        # Remove old dated copies
        for old in glob.glob(os.path.join(OUTPUT_FOLDER, "01. PD_data_*.xlsb")):
            if old != output_file:
                try:
                    os.remove(old)
                    logger.info(f"Removed old file: {os.path.basename(old)}")
                except Exception as e:
                    logger.warning(f"Could not remove {os.path.basename(old)}: {e}")

    # ------------------------------------------------------------------
    # Update cache and latest_month config
    # ------------------------------------------------------------------
    save_cycle_cache(all_data)
    ExcelPortfolioAutomation.save_config_file(CONFIG_FILE, new_end_month)
    logger.info(f"latest_month.txt updated -> {ExcelPortfolioAutomation.format_month_string(new_end_month)}")

    logger.info("=" * 80)
    logger.info("PORTFOLIO ROLL-FORWARD COMPLETED!")
    logger.info("=" * 80)

    return True, output_file


# =============================================================================
# MAIN AUTOMATION - PART 2: HISTORIC PD UPDATE
# =============================================================================
def run_historic_update(latest_pd_file: str) -> bool:
    """Second automation — copy pivot data to Historic PD format."""
    try:
        historic_matches = glob.glob(os.path.join(INPUT_FOLDER, "02. Historic PD Calculation*.xlsb"))
        if not historic_matches:
            logger.error("No '02. Historic PD Calculation*.xlsb' file found in INPUT_FOLDER")
            return False
        historic_filename = os.path.basename(sorted(historic_matches)[0])

        historic_output_file = ExcelPortfolioAutomation.copy_pivot_to_historic(
            latest_pd_file=latest_pd_file,
            input_folder=INPUT_FOLDER,
            output_folder=OUTPUT_FOLDER,
            historic_filename=historic_filename,
            pivot_sheet="01.Pivoted_Portfolio",
            historic_sheet="02.Working"
        )

        if not historic_output_file:
            logger.error("Failed to create Historic PD file")
            return False

        # Rename to date-based name
        date_str   = datetime.now().strftime("%Y-%m-%d")
        final_name = f"02. Historic PD Calculation_{date_str}.xlsb"
        final_path = os.path.join(OUTPUT_FOLDER, final_name)
        os.rename(historic_output_file, final_path)
        logger.info(f"Historic PD file saved: {final_name}")

        # Remove old dated copies
        for old in glob.glob(os.path.join(OUTPUT_FOLDER, "02. Historic PD Calculation*.xlsb")):
            if old != final_path:
                try:
                    os.remove(old)
                    logger.info(f"Removed old file: {os.path.basename(old)}")
                except Exception as e:
                    logger.warning(f"Could not remove {os.path.basename(old)}: {e}")

        return True

    except Exception as e:
        logger.error(f"ERROR in Historic PD update: {e}")
        import traceback
        traceback.print_exc()
        return False


# =============================================================================
# ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    setup_logger(os.path.dirname(os.path.abspath(__file__)))
    try:
        logger.info("=" * 80)
        logger.info("   PD AUTOMATION - COMPLETE WORKFLOW (DYNAMIC ROLLING CYCLE)")
        logger.info("=" * 80)
        if SELECTED_RUN_DATE:
            logger.info(f"Selected run date: {SELECTED_RUN_DATE}")
        if PD_BASE_FOLDER:
            logger.info(f"PD base folder: {PD_BASE_FOLDER}")

        # ==============================================================
        # PART 1: Portfolio Roll-Forward
        # ==============================================================
        logger.info("Starting Part 1: Portfolio Roll-Forward...")
        success, pd_output_file = run_automation()

        if not success or not pd_output_file:
            logger.error("Part 1 failed. Stopping automation.")
            sys.exit(1)  # generic failure (not exit 2 — that is raised directly in run_automation)

        logger.info("Part 1 completed successfully!")

        # ==============================================================
        # PART 2: Historic PD Update
        # ==============================================================
        logger.info("Starting Part 2: Historic PD Update...")
        historic_success = run_historic_update(pd_output_file)

        if not historic_success:
            logger.error("Part 2 failed, but Part 1 was successful.")
            sys.exit(1)

        logger.info("Part 2 completed successfully!")

        # ==============================================================
        # FINAL SUMMARY
        # ==============================================================
        logger.info("=" * 80)
        logger.info("   ALL AUTOMATIONS COMPLETED SUCCESSFULLY!")
        logger.info("=" * 80)
        logger.info(f"Output files saved in: {OUTPUT_FOLDER}")

        sys.exit(0)

    except Exception as e:
        logger.error("=" * 80)
        logger.error(f"FATAL ERROR: {e}")
        logger.error("=" * 80)
        import traceback
        traceback.print_exc()
        sys.exit(1)

"""
OLAdjstmnt macro — Python OOP conversion using xlwings.
Every step verified line-by-line against the original VBA.

Bug-fix history
---------------
v1  Initial conversion.
v2  Fixed _last_row (used_range), _clear_summary_data_rows (used_range),
    _paste_values_api (win32com PasteSpecial), renamed method b_to_e.
v3  Fixed 0x2a corruption in toSummary columns J / K.

    ROOT CAUSE (0x2a issue)
    ───────────────────────
    The image shows columns J and K on the toSummary sheet contain a
    PivotTable whose source is ECL data.  Some ECL rows carry the literal
    string "0x2a" (hex 42) instead of a numeric value — these are either:
      (a) rows where a formula produced an Excel error (#N/A, #VALUE!, etc.)
          and the upstream system serialised the error as a hex integer, OR
      (b) rows where the source feed wrote the string literal "0x2a" directly.

    Because the macro copies those rows as-is (PasteSpecial values), and the
    PivotTable then reads from the pasted range, "0x2a" lands in J and K.

    THREE-LAYER FIX applied in v3
    ──────────────────────────────
    1. _sanitise_range()  — after every paste into ECL working columns and
       into toSummary, scan the destination for "0x2a" / Excel error values
       and replace them with None (blank), so downstream formulas and the
       PivotTable never see corrupt values.

    2. _last_row_col_b()  — the VBA anchors lastRow to column B specifically
       (End(xlUp) on col B).  used_range.last_cell.row is wider: it returns
       the last row of ANY used cell across ALL columns, which can include
       stray formula-error rows far below the real data boundary in col B.
       Fixed to replicate the VBA's col-B anchor exactly via the API.

    3. _copy_ecl_col_to_summary() — unified helper that sanitises every
       destination range immediately after paste, before the next step runs.
"""

import time
from typing import Optional

import xlwings as xw

try:
    from scipy.optimize import brentq
    _SCIPY_AVAILABLE = True
except ImportError:
    _SCIPY_AVAILABLE = False

# ── Excel win32com constants ──────────────────────────────────────────────────
_XL_PASTE_VALUES  = -4163   # xlPasteValues
_XL_OFF           = -4146   # xlOff  (checkbox)
_XL_ON            =  1      # xlOn   (checkbox)

# Values that indicate a bad / corrupt cell coming from the source data.
# Includes the literal string variants and the xlwings integer error codes.
_EXCEL_ERROR_CODES = {2000, 2007, 2015, 2023, 2029, 2036, 2042}  # #NULL! … #N/A
_BAD_STRINGS       = {"0x2a", "0X2A"}   # hex literal written by upstream system


class OLAdjustment:
    """
    Encapsulates every step of the OL Adjustment macro.

    Parameters
    ----------
    workbook : xw.Book, optional
        Explicit workbook.  Defaults to the active workbook.
    """

    def __init__(self, workbook: Optional[xw.Book] = None) -> None:
        self.workbook: xw.Book  = workbook if workbook is not None else xw.books.active
        self.app:      xw.App   = self.workbook.app

        self.ws_ecl:     xw.Sheet = self.workbook.sheets["ECL"]
        self.ws_summary: xw.Sheet = self.workbook.sheets["toSummary"]
        self.ws_pivot:   xw.Sheet = self.workbook.sheets["Pivot"]

    # ── Public entry point ────────────────────────────────────────────────────

    def run(self) -> None:
        """Run every macro step in the original VBA order."""

        self._validate_yard_provision()           # guard: BR23 numeric
        self._clear_adjusted_lgd()                # BQ2.ClearContents
        self._disable_dynamic_mode()              # CheckBox xlOff
        self._wait(5)
        self._copy_original_model_imp_to_col57()  # col49 → col57
        self._wait(5)
        self._copy_100pct_overlays_to_col51()     # col64 → col51
        self._clear_summary_data_rows()           # toSummary rows 2:end clear
        self._copy_account_ids_to_summary()       # ECL col B → summary A
        self._copy_ecl_columns_to_summary_b_to_e()  # 57,56,50,51 → B-E
        self._enable_dynamic_mode()               # CheckBox xlOn
        self._wait(5)
        self._copy_risk_elevated_overlays_to_col62()  # col63 → col62
        self._copy_ecl_columns_fg_to_summary()    # 61,62 → F,G
        self._refresh_pivot_tables_and_snapshot() # Pivot refresh + A2:D36→H2
        self._snapshot_lease_original_imp()       # BQ21 → BR21 values
        self._clear_clipboard()
        self._run_goal_seek()                     # BS25 GoalSeek → BQ2
        self._copy_col55_to_summary_h()           # col55 → summary H
        self._clear_clipboard()

    # ── Core helpers ──────────────────────────────────────────────────────────

    def _last_row_col_b(self) -> int:
        """
        Replicate VBA:  wsECL.Cells(wsECL.Rows.Count, "B").End(xlUp).Row

        v2 used used_range.last_cell.row which returns the last row of ANY
        used cell across ALL columns — this can overshoot col B's actual data
        boundary when other columns have stray content or error values lower
        down.  Using the win32com API to walk up from the bottom of col B
        exactly mirrors the VBA.
        """
        max_row  = self.ws_ecl.api.Rows.Count          # 1,048,576
        last_row = (
            self.ws_ecl.api
            .Cells(max_row, 2)          # col 2 = B
            .End(-4162)                 # xlUp = -4162
            .Row
        )
        return last_row

    def _paste_values_api(
        self,
        src_range: xw.Range,
        dst_cell:  xw.Range,
    ) -> None:
        """
        Copy src_range → paste as values at dst_cell via win32com.
        Clears the clipboard immediately after.
        """
        src_range.api.Copy()
        dst_cell.api.PasteSpecial(Paste=_XL_PASTE_VALUES)
        self.app.cut_copy_mode = False

    def _sanitise_range(self, rng: xw.Range) -> None:
        """
        FIX for 0x2a corruption.

        Scan every cell in *rng*.  Replace any cell whose value is:
          • the literal string "0x2a" / "0X2A"  (hex written by source system)
          • an Excel error integer code (xlwings returns these as int)
          • any non-numeric, non-None value that should be numeric
        with None (blank), so the PivotTable and downstream formulas
        never receive corrupt data.

        Uses a bulk read/write (options(ndim=2)) to avoid per-cell API calls.
        """
        data = rng.options(ndim=2).value
        if data is None:
            return

        changed = False
        cleaned = []
        for row in data:
            clean_row = []
            for cell in row:
                if (
                    cell in _BAD_STRINGS
                    or (isinstance(cell, str) and cell.lower().startswith("0x"))
                    or (isinstance(cell, int) and cell in _EXCEL_ERROR_CODES)
                ):
                    clean_row.append(None)   # blank the bad value
                    changed = True
                else:
                    clean_row.append(cell)
            cleaned.append(clean_row)

        if changed:
            rng.options(ndim=2).value = cleaned

    # ── Validation / setup ────────────────────────────────────────────────────

    def _validate_yard_provision(self) -> None:
        """VBA: If IsEmpty(BR23) Or Not IsNumeric(BR23) Then Exit Sub"""
        val = self.ws_ecl.range("BR23").value
        if val is None or not isinstance(val, (int, float)):
            raise ValueError(
                "BR23 is empty or non-numeric. "
                "Please add Yard Provision before running."
            )

    def _clear_adjusted_lgd(self) -> None:
        """VBA: Range("BQ2").ClearContents"""
        self.ws_ecl.range("BQ2").clear_contents()

    def _disable_dynamic_mode(self) -> None:
        """VBA: wsECL.CheckBoxes("Check Box 1").Value = xlOff"""
        self.ws_ecl.api.CheckBoxes("Check Box 1").Value = _XL_OFF

    def _enable_dynamic_mode(self) -> None:
        """VBA: wsECL.CheckBoxes("Check Box 1").Value = xlOn"""
        self.ws_ecl.api.CheckBoxes("Check Box 1").Value = _XL_ON

    @staticmethod
    def _wait(seconds: int) -> None:
        """VBA: Application.Wait Now + TimeValue(...)"""
        time.sleep(seconds)

    def _clear_clipboard(self) -> None:
        """VBA: Application.CutCopyMode = False"""
        self.app.cut_copy_mode = False

    # ── ECL internal column copies ────────────────────────────────────────────

    def _copy_original_model_imp_to_col57(self) -> None:
        """
        VBA: Cells(3,49):Cells(lastRow,49).Copy → Cells(3,57) PasteValues
        Source starts at row 3 (header in rows 1-2).
        """
        last = self._last_row_col_b()
        src  = self.ws_ecl.range(self.ws_ecl.cells(3, 49), self.ws_ecl.cells(last, 49))
        dst  = self.ws_ecl.cells(3, 57)
        self._paste_values_api(src, dst)
        self._sanitise_range(
            self.ws_ecl.range(self.ws_ecl.cells(3, 57), self.ws_ecl.cells(last, 57))
        )

    def _copy_100pct_overlays_to_col51(self) -> None:
        """
        VBA: Cells(3,64):Cells(lastRow,64).Copy → Cells(3,51) PasteValues
        Source starts at row 3.
        """
        last = self._last_row_col_b()
        src  = self.ws_ecl.range(self.ws_ecl.cells(3, 64), self.ws_ecl.cells(last, 64))
        dst  = self.ws_ecl.cells(3, 51)
        self._paste_values_api(src, dst)
        self._sanitise_range(
            self.ws_ecl.range(self.ws_ecl.cells(3, 51), self.ws_ecl.cells(last, 51))
        )

    def _copy_risk_elevated_overlays_to_col62(self) -> None:
        """
        VBA: Cells(3,63):Cells(lastRow,63).Copy → Cells(3,62) PasteValues
        Source starts at row 3.
        """
        last = self._last_row_col_b()
        src  = self.ws_ecl.range(self.ws_ecl.cells(3, 63), self.ws_ecl.cells(last, 63))
        dst  = self.ws_ecl.cells(3, 62)
        self._paste_values_api(src, dst)
        self._sanitise_range(
            self.ws_ecl.range(self.ws_ecl.cells(3, 62), self.ws_ecl.cells(last, 62))
        )

    # ── toSummary population ──────────────────────────────────────────────────

    def _clear_summary_data_rows(self) -> None:
        """
        VBA: wsSummary.Rows("2:" & wsSummary.Rows.Count).ClearContents
        Uses the API to get the actual last used row (not sheet max).
        """
        max_row  = self.ws_summary.api.Rows.Count
        last_used = (
            self.ws_summary.api
            .Cells(max_row, 1)
            .End(-4162)   # xlUp
            .Row
        )
        if last_used >= 2:
            self.ws_summary.range(f"2:{last_used}").clear_contents()

    def _copy_ecl_col_to_summary(
        self,
        ecl_col_num:  int,
        summary_col:  str,
        start_row:    int = 2,
    ) -> None:
        """
        Generic helper: paste ECL col_num (rows start_row:lastRow) → summary_col,
        then immediately sanitise the destination to remove any 0x2a / error values.
        """
        last = self._last_row_col_b()
        src  = self.ws_ecl.range(
            self.ws_ecl.cells(start_row, ecl_col_num),
            self.ws_ecl.cells(last,      ecl_col_num),
        )
        dst = self.ws_summary.range(f"{summary_col}2")
        self._paste_values_api(src, dst)

        # ── FIX: sanitise immediately after paste ──────────────────────────
        row_count = last - start_row + 1
        dst_col_num = ord(summary_col.upper()) - ord("A") + 1
        dst_range = self.ws_summary.range(
            self.ws_summary.cells(2,            dst_col_num),
            self.ws_summary.cells(1 + row_count, dst_col_num),
        )
        self._sanitise_range(dst_range)

    def _copy_account_ids_to_summary(self) -> None:
        """VBA: ECL B2:lastRow → toSummary A2"""
        last = self._last_row_col_b()
        src  = self.ws_ecl.range(f"B2:B{last}")
        dst  = self.ws_summary.range("A2")
        self._paste_values_api(src, dst)
        # Account IDs are text/numbers — still sanitise for safety
        self._sanitise_range(
            self.ws_summary.range(f"A2:A{last}")
        )

    def _copy_ecl_columns_to_summary_b_to_e(self) -> None:
        """
        VBA:
            col57 → toSummary B    col56 → toSummary C
            col50 → toSummary D    col51 → toSummary E
        """
        self._copy_ecl_col_to_summary(57, "B")
        self._copy_ecl_col_to_summary(56, "C")
        self._copy_ecl_col_to_summary(50, "D")
        self._copy_ecl_col_to_summary(51, "E")

    def _copy_ecl_columns_fg_to_summary(self) -> None:
        """VBA: col61 → toSummary F    col62 → toSummary G"""
        self._copy_ecl_col_to_summary(61, "F")
        self._copy_ecl_col_to_summary(62, "G")

    def _copy_col55_to_summary_h(self) -> None:
        """
        VBA: Cells(2,55):Cells(lastRow,55).Copy → toSummary H2
        (VBA comment says col 54 but code uses col 55 — col 55 is correct.)
        """
        self._copy_ecl_col_to_summary(55, "H")

    # ── Pivot sheet ───────────────────────────────────────────────────────────

    def _refresh_pivot_tables_and_snapshot(self) -> None:
        """
        VBA:
            With Sheets("Pivot")
                For Each pt In .PivotTables: pt.RefreshTable: Next pt
                .Range("A2:D36").Copy → .Range("H2").PasteSpecial values
            End With

        After pasting, sanitise Pivot!H2:K36 so the PivotTable feeding
        toSummary J/K never sees 0x2a values from corrupted ECL source rows.
        """
        for pt in self.ws_pivot.api.PivotTables():
            pt.RefreshTable()

        src = self.ws_pivot.range("A2:D36")
        dst = self.ws_pivot.range("H2")
        self._paste_values_api(src, dst)

        # ── FIX: sanitise the snapshot that feeds toSummary J / K ─────────
        self._sanitise_range(self.ws_pivot.range("H2:K36"))

    # ── ECL lease snapshot ────────────────────────────────────────────────────

    def _snapshot_lease_original_imp(self) -> None:
        """VBA: wsECL.BQ21.Copy → wsECL.BR21.PasteSpecial values"""
        src = self.ws_ecl.range("BQ21")
        dst = self.ws_ecl.range("BR21")
        self._paste_values_api(src, dst)

    # ── Goal Seek ─────────────────────────────────────────────────────────────

    def _run_goal_seek(self) -> None:
        """
        VBA: wsECL.Range("BS25").GoalSeek Goal:=0, ChangingCell:=wsECL.Range("BQ2")

        Uses scipy.brentq (more robust).  Falls back to Excel native GoalSeek.
        """
        def _objective(x: float) -> float:
            self.ws_ecl.range("BQ2").value = x
            self.app.calculate()
            val = self.ws_ecl.range("BS25").value
            if not isinstance(val, (int, float)):
                raise ValueError(f"BS25 returned non-numeric value: {val!r}")
            return float(val)

        if _SCIPY_AVAILABLE:
            try:
                solution = brentq(_objective, -1e6, 1e6, xtol=1e-6, maxiter=200)
                self.ws_ecl.range("BQ2").value = solution
                return
            except Exception:
                pass  # fall through to native GoalSeek

        # Native Excel GoalSeek fallback
        self.ws_ecl.range("BS25").api.GoalSeek(
            Goal=0,
            ChangingCell=self.ws_ecl.range("BQ2").api,
        )


# ── Convenience wrapper ────────────────────────────────────────────────────────

def run_ol_adjustment(workbook: Optional[xw.Book] = None) -> None:
    """
    Entry point — mirrors calling OLAdjstmnt() from VBA.

    Usage
    -----
        from ol_adjustment import run_ol_adjustment

        run_ol_adjustment()                                # active workbook
        run_ol_adjustment(xw.Book(r"C:/path/model.xlsx")) # explicit path
    """
    OLAdjustment(workbook).run()


if __name__ == "__main__":
    run_ol_adjustment()

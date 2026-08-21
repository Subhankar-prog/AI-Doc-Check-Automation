"""
input_tracker.py — Live status tracking written back into input_map.xlsx

As each file is processed, this module writes the result status
directly into the input Excel sheet, Column C ("Run Status"):

    PENDING   → not yet processed (default)
    RUNNING   → currently being processed
    SUCCESS   → done, extracted OK
    FAILED    → done, engine returned failure
    ERROR     → script exception occurred
    TIMEOUT   → engine did not respond in time
    SKIPPED   → already done in a previous run (won't re-run)

This way, even if the script crashes at file 47 of 100:
  - Rows 1-46 show SUCCESS/FAILED
  - Row 47 shows RUNNING (was in progress)
  - Rows 48-100 show PENDING

On the NEXT run, use:
    python main.py --resume
to skip all SKIPPED/SUCCESS/FAILED rows and continue from PENDING.
"""
import logging
from pathlib import Path
from typing import List, Dict, Any

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

logger = logging.getLogger(__name__)

# Column indices in each sheet (1-based)
COL_FILE_PATH  = 1   # A - file path (user fills this)
COL_NOTES      = 2   # B - notes (optional)
COL_STATUS     = 3   # C - Run Status (written by automation)
COL_RESULT_REF = 4   # D - Output row reference (e.g. "See output row 5")

# Status values
STATUS_PENDING  = "PENDING"
STATUS_RUNNING  = "RUNNING"
STATUS_SUCCESS  = "SUCCESS"
STATUS_FAILED   = "FAILED"
STATUS_ERROR    = "ERROR"
STATUS_TIMEOUT  = "TIMEOUT"
STATUS_SKIPPED  = "SKIPPED"

# Colours for each status
STATUS_FILLS = {
    STATUS_PENDING:  PatternFill("solid", fgColor="F5F5F5"),  # light grey
    STATUS_RUNNING:  PatternFill("solid", fgColor="FFF9C4"),  # yellow
    STATUS_SUCCESS:  PatternFill("solid", fgColor="C8E6C9"),  # green
    STATUS_FAILED:   PatternFill("solid", fgColor="FFCDD2"),  # red
    STATUS_ERROR:    PatternFill("solid", fgColor="FFE0B2"),  # orange
    STATUS_TIMEOUT:  PatternFill("solid", fgColor="E1BEE7"),  # purple
    STATUS_SKIPPED:  PatternFill("solid", fgColor="E0E0E0"),  # grey
}
STATUS_FONTS = {
    STATUS_PENDING:  Font(name="Calibri", size=10, color="888888", italic=True),
    STATUS_RUNNING:  Font(name="Calibri", size=10, color="F57F17", bold=True),
    STATUS_SUCCESS:  Font(name="Calibri", size=10, color="1B5E20", bold=True),
    STATUS_FAILED:   Font(name="Calibri", size=10, color="B71C1C", bold=True),
    STATUS_ERROR:    Font(name="Calibri", size=10, color="E65100", bold=True),
    STATUS_TIMEOUT:  Font(name="Calibri", size=10, color="4A148C", bold=True),
    STATUS_SKIPPED:  Font(name="Calibri", size=10, color="616161", italic=True),
}


class InputTracker:
    """
    Tracks processing state for each row in the input Excel.
    Writes status back to Column C of the input sheet live.
    """

    def __init__(self, input_file: Path):
        self.input_file = input_file
        self._wb = openpyxl.load_workbook(str(input_file))
        # Map: (sheet_name, row_number) -> current status
        self._row_map: Dict[str, Dict[int, Any]] = {}
        self._ensure_status_column()

    # ── Setup ─────────────────────────────────────────────
    def _ensure_status_column(self):
        """Add 'Run Status' and 'Output Ref' header to col C & D if not present."""
        for sheet_name in self._wb.sheetnames:
            ws = self._wb[sheet_name]
            # Row 2 = headers row
            if ws.cell(row=2, column=COL_STATUS).value != "Run Status":
                c = ws.cell(row=2, column=COL_STATUS, value="Run Status")
                c.fill      = PatternFill("solid", fgColor="1E3A5F")
                c.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
                c.alignment = Alignment(horizontal="center", vertical="center")
            if ws.cell(row=2, column=COL_RESULT_REF).value != "Output Row":
                d = ws.cell(row=2, column=COL_RESULT_REF, value="Output Row")
                d.fill      = PatternFill("solid", fgColor="1E3A5F")
                d.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
                d.alignment = Alignment(horizontal="center", vertical="center")
            # Set column widths
            ws.column_dimensions["C"].width = 14
            ws.column_dimensions["D"].width = 14
        self._save()

    # ── Load rows ──────────────────────────────────────────
    def load_records(self, doc_type_filter: str = None,
                     skip_done: bool = False) -> List[Dict]:
        """
        Read all records from the input Excel.
        Returns list of dicts with added 'row_num' and 'sheet_name' keys.

        Args:
            doc_type_filter: if set, only load rows from this sheet
            skip_done:       if True, skip rows already marked SUCCESS/FAILED/SKIPPED
        """
        from config import DOCUMENT_TYPES
        records = []

        for sheet_name in self._wb.sheetnames:
            if sheet_name not in DOCUMENT_TYPES:
                continue
            if doc_type_filter and sheet_name.lower() != doc_type_filter.lower():
                # Try partial match
                if doc_type_filter.lower() not in sheet_name.lower():
                    continue

            ws = self._wb[sheet_name]
            for row_num in range(3, ws.max_row + 1):
                file_path = ws.cell(row=row_num, column=COL_FILE_PATH).value
                if not file_path:
                    continue
                file_path = str(file_path).strip()
                if not file_path or "<" in file_path or ">" in file_path or file_path == "None":
                    continue

                current_status = ws.cell(row=row_num, column=COL_STATUS).value or STATUS_PENDING

                # Skip already-successful rows if --resume
                # (PENDING, FAILED, ERROR, TIMEOUT, and RUNNING files will still be executed)
                if skip_done and current_status in (STATUS_SUCCESS, STATUS_SKIPPED):
                    logger.info(f"  SKIPPING (already {current_status}): {file_path}")
                    continue



                records.append({
                    "file_path":   file_path,
                    "doc_type":    sheet_name,
                    "row_num":     row_num,
                    "sheet_name":  sheet_name,
                    "prev_status": current_status,
                })

        return records

    # ── Status updates ────────────────────────────────────
    def mark_running(self, sheet_name: str, row_num: int):
        """Mark a row as currently being processed."""
        self._write_status(sheet_name, row_num, STATUS_RUNNING, "In progress...")

    def mark_done(self, sheet_name: str, row_num: int,
                  status: str, output_row: int = None):
        """
        Mark a row as completed.
        status: one of SUCCESS / FAILED / ERROR / TIMEOUT
        output_row: row number in the output Excel for cross-reference
        """
        ref = f"Row {output_row}" if output_row else ""
        self._write_status(sheet_name, row_num, status, ref)

    def _write_status(self, sheet_name: str, row_num: int,
                      status: str, ref: str = ""):
        """Write status and ref to cols C & D in the input Excel."""
        try:
            ws   = self._wb[sheet_name]
            cell = ws.cell(row=row_num, column=COL_STATUS, value=status)
            cell.fill      = STATUS_FILLS.get(status, PatternFill())
            cell.font      = STATUS_FONTS.get(status, Font(name="Calibri", size=10))
            cell.alignment = Alignment(horizontal="center", vertical="center")

            ref_cell = ws.cell(row=row_num, column=COL_RESULT_REF, value=ref)
            ref_cell.alignment = Alignment(horizontal="center", vertical="center")
            ref_cell.font      = Font(name="Calibri", size=9, color="555555")

            self._save()
            logger.debug(f"Input tracker: [{sheet_name}] row {row_num} -> {status}")
        except Exception as e:
            logger.warning(f"Could not write status to input Excel: {e}")

    def _save(self):
        try:
            self._wb.save(str(self.input_file))
        except PermissionError:
            logger.warning(
                "Could not save input_map.xlsx — it may be open in Excel. "
                "Close the file and the status will update on next write."
            )

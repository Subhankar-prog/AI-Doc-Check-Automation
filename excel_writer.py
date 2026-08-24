"""
excel_writer.py - Write automation results to Excel with one tab per document type

Output Excel structure:
    Sheet "Aadhaar Card"           <- all Aadhaar Card results (100s of rows)
    Sheet "PAN Card"               <- all PAN Card results
    Sheet "Graduation Certificate" <- all graduation cert results
    ... (one sheet per document type, created on demand)

Each sheet:
    - Fixed columns first: Filename, Document Type, Expected Status, Actual Status, Final Output, Error Message
    - Followed by dynamic extracted field columns (vary per doc type)
    - Green rows = PASS, Red rows = FAIL, Yellow = REVIEW
    - Headers auto-generated on first row written to each sheet
    - Saved after every row (crash-safe)
    - A "Summary" sheet is written at the end with counts & timing per doc type
"""
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import OUTPUT_DIR, OUTPUT_EXCEL_PREFIX, FIXED_COLUMNS, DOCUMENT_TYPES, KNOWN_VALIDATION_REJECTION_PATTERNS

logger = logging.getLogger(__name__)

# ── Color palette ─────────────────────────────────────────────────────────────
SUCCESS_FILL = PatternFill("solid", fgColor="D6F4D2")   # light green
FAILED_FILL  = PatternFill("solid", fgColor="FAD4D4")   # light red
UNKNOWN_FILL = PatternFill("solid", fgColor="FFF3CC")   # light yellow
HEADER_FILL  = PatternFill("solid", fgColor="1E3A5F")   # dark navy
HEADER_FONT  = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
CELL_FONT    = Font(name="Calibri", size=10)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_ALIGN   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
THIN         = Side(style="thin")
THIN_BORDER  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Excel sheet name max 31 chars
_MAX_SHEET_NAME = 31


def _safe_sheet_name(name: str) -> str:
    """Truncate and sanitize sheet name for Excel."""
    invalid = r'\/:*?[]'
    for ch in invalid:
        name = name.replace(ch, "_")
    return name[:_MAX_SHEET_NAME]


def classify_test_result(status: str, error_message: str, pre_check_reject: bool = False) -> tuple:
    """
    Decide whether this row was an expected outcome or a genuine problem,
    based on the actual Status, the portal's own error message, AND (if
    available) the pre-upload size/format prediction.

    Logic:
      - If pre-upload check predicted rejection (oversized / wrong format),
        OR error message matches KNOWN_VALIDATION_REJECTION_PATTERNS:
        -> Expected Status = REJECT
      - Otherwise:
        -> Expected Status = SUCCESS

    Then Final Output = PASS if actual matches expected,
         FAIL if it doesn't,
         REVIEW if actual outcome couldn't be determined.

    Returns:
        (expected_status, final_output)
    """
    status_upper = (status or "").upper()
    err_lower    = (error_message or "").lower()

    is_known_validation_failure = any(
        p in err_lower for p in KNOWN_VALIDATION_REJECTION_PATTERNS
    )
    expected_status = "REJECT" if (is_known_validation_failure or pre_check_reject) else "SUCCESS"

    if status_upper == "SUCCESS":
        actual_outcome = "SUCCESS"
    elif status_upper in ("FAILED", "ERROR", "TIMEOUT"):
        actual_outcome = "REJECT"
    else:
        actual_outcome = "UNKNOWN"

    if actual_outcome == "UNKNOWN":
        final_output = "REVIEW"
    elif expected_status == actual_outcome:
        final_output = "PASS"
    else:
        final_output = "FAIL"

    return expected_status, final_output


class ExcelWriter:
    """
    Manages one Excel workbook with a separate worksheet per document type.
    Usage:
        writer = ExcelWriter()
        writer.append_row(result_dict)   # call after each document
        writer.finalize()                # call at the end to write Summary tab
    """

    def __init__(self, resume_path: str = None):
        """
        If resume_path is given and the file exists, reopen it and continue
        appending to it (same sheets, same Summary later) instead of always
        creating a brand-new timestamped file. This lets a stopped run pick
        back up into ONE combined output file rather than splitting results
        across multiple files.
        """
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        # Per-sheet state: sheet_name -> {"ws": worksheet, "headers": [...], "row_count": int}
        self._sheets: Dict[str, dict] = {}

        # Per-doc-type stats for Summary sheet (SUCCESS/FAILED derived directly from Final Output PASS/FAIL)
        self._stats: Dict[str, Dict[str, Any]] = {}

        if resume_path and Path(resume_path).exists():
            self.filepath = Path(resume_path)
            self.wb = openpyxl.load_workbook(self.filepath)
            self._load_existing_workbook()
            logger.info(f"Resuming into existing output Excel: {self.filepath}")
            logger.info(
                "Note: total/average processing-time stats for ALREADY-written "
                "rows could not be recovered (that value isn't stored per-row) "
                "— time stats in the Summary will only reflect this session's "
                "new rows going forward. Success/Failed/Error counts are "
                "fully accurate."
            )
        else:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.filepath = OUTPUT_DIR / f"{OUTPUT_EXCEL_PREFIX}_{ts}.xlsx"
            self.wb = openpyxl.Workbook()
            self.wb.remove(self.wb.active)   # remove the default empty sheet
            logger.info(f"Output Excel: {self.filepath}")

    def _load_existing_workbook(self):
        """
        Rebuild internal sheet/stats state by reading an already-existing
        output workbook, so appending new rows continues correctly (right
        headers, right next row number, right Summary counts).
        """
        for ws in self.wb.worksheets:
            if ws.title == "Summary":
                continue   # Summary is fully regenerated by finalize()

            headers = [c.value for c in ws[1] if c.value is not None]
            if not headers:
                continue
            row_count = max(0, ws.max_row - 1)   # minus header row
            self._sheets[ws.title] = {"ws": ws, "headers": headers, "row_count": row_count}

            # Rebuild stats for the Summary sheet from what's already on disk.
            # Success/Failed/Error are derived from the Final Output column,
            # same as append_row() does for new rows.
            try:
                dt_idx = headers.index("Document Type") + 1
                fo_idx = headers.index("Final Output") + 1
            except ValueError:
                continue   # older-format sheet, skip stat rebuild for it

            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                doc_type = str(row[dt_idx - 1].value or "Unknown")
                final    = str(row[fo_idx - 1].value or "REVIEW").upper()

                if doc_type not in self._stats:
                    self._stats[doc_type] = {
                        "SUCCESS": 0, "FAILED": 0, "ERROR": 0, "TOTAL": 0, "TOTAL_TIME": 0.0
                    }
                self._stats[doc_type]["TOTAL"] += 1
                if final == "PASS":
                    self._stats[doc_type]["SUCCESS"] += 1
                elif final == "FAIL":
                    self._stats[doc_type]["FAILED"] += 1
                else:
                    self._stats[doc_type]["ERROR"] += 1

    # ── Public API ────────────────────────────────────────────
    def append_row(self, row_dict: Dict[str, Any]):
        """
        Append a result row to the sheet for its document type.
        Creates the sheet and header row automatically on first use.
        Saves the workbook immediately after writing (crash-safe).
        """
        doc_type   = str(row_dict.get("Document Type", "Unknown")).strip()
        sheet_name = _safe_sheet_name(doc_type) if doc_type else "Unknown"

        # Derive Expected Status / Final Output from the actual Status + Error Message
        raw_status = str(row_dict.get("Status", "UNKNOWN"))
        error_msg  = str(row_dict.get("Error Message", ""))
        pre_check_reject = bool(row_dict.get("_pre_check_reject", False))
        expected_status, final_output = classify_test_result(raw_status, error_msg, pre_check_reject)
        row_dict["Expected Status"] = expected_status
        row_dict["Actual Status"]   = raw_status
        row_dict["Final Output"]    = final_output

        # Create sheet on first use for this doc type
        if sheet_name not in self._sheets:
            self._create_sheet(sheet_name, row_dict)

        sheet_state = self._sheets[sheet_name]
        ws          = sheet_state["ws"]

        # Check for new dynamic columns (ignoring private keys starting with '_' and raw 'Status')
        existing = sheet_state["headers"]
        new_cols  = [k for k in row_dict if k not in existing and not k.startswith("_") and k != "Status"]
        for col in new_cols:
            existing.append(col)
            col_idx = len(existing)
            self._write_header_cell(ws, 1, col_idx, col)

        # Determine row fill colour based on the Final Output verdict (PASS/FAIL)
        fill = (SUCCESS_FILL if final_output == "PASS"
                else FAILED_FILL if final_output == "FAIL"
                else UNKNOWN_FILL)

        # Write the data row
        row_num = sheet_state["row_count"] + 2   # +2 because row 1 = header
        for col_idx, header in enumerate(existing, start=1):
            value = row_dict.get(header, "")
            cell  = ws.cell(row=row_num, column=col_idx, value=value)
            cell.fill      = fill
            cell.font      = CELL_FONT
            cell.alignment = CENTER_ALIGN if col_idx <= 5 else LEFT_ALIGN
            cell.border    = THIN_BORDER

        ws.row_dimensions[row_num].height = 18
        sheet_state["row_count"] += 1

        # Track stats and timing for Summary sheet
        ptime = 0.0
        try:
            ptime = float(row_dict.get("_processing_time", row_dict.get("Processing Time (s)", 0.0)))
        except (ValueError, TypeError):
            ptime = 0.0

        if doc_type not in self._stats:
            self._stats[doc_type] = {
                "SUCCESS": 0, "FAILED": 0, "ERROR": 0, "TOTAL": 0, "TOTAL_TIME": 0.0
            }
        self._stats[doc_type]["TOTAL"] += 1
        self._stats[doc_type]["TOTAL_TIME"] += ptime

        # Success & Failed are fetched directly from Final Output logic:
        # PASS -> Success
        # FAIL -> Failed
        # REVIEW/ERROR -> Error
        if final_output == "PASS":
            self._stats[doc_type]["SUCCESS"] += 1
        elif final_output == "FAIL":
            self._stats[doc_type]["FAILED"] += 1
        else:
            self._stats[doc_type]["ERROR"] += 1

        # Auto-fit column widths and save
        self._autofit(ws, existing)
        self._save()

    def finalize(self) -> str:
        """
        Write the Summary sheet (overall stats and timing per doc type) and save.
        Returns the output file path.
        """
        self._write_summary_sheet()
        # Freeze header row on every sheet
        for state in self._sheets.values():
            state["ws"].freeze_panes = "A2"
        self._save()
        logger.info(f"Excel saved: {self.filepath}")
        return str(self.filepath)

    # ── Internal helpers ──────────────────────────────────────
    def _create_sheet(self, sheet_name: str, first_row: Dict[str, Any]):
        """Create a new worksheet with header row for a document type."""
        ws = self.wb.create_sheet(title=sheet_name)

        # Build header list: fixed columns first, then any extra dynamic cols (no private '_')
        dynamic = [k for k in first_row if k not in FIXED_COLUMNS and not k.startswith("_") and k != "Status"]
        headers = FIXED_COLUMNS + dynamic

        for col_idx, h in enumerate(headers, start=1):
            self._write_header_cell(ws, 1, col_idx, h)

        ws.row_dimensions[1].height = 22
        ws.freeze_panes = "A2"

        self._sheets[sheet_name] = {
            "ws":        ws,
            "headers":   headers,
            "row_count": 0,
        }
        logger.info(f"Created sheet: '{sheet_name}'  headers: {headers}")

    def _write_header_cell(self, ws, row: int, col: int, text: str):
        cell = ws.cell(row=row, column=col, value=text)
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = CENTER_ALIGN
        cell.border    = THIN_BORDER

    def _autofit(self, ws, headers: List[str]):
        """Set column widths based on content (capped at 50 chars)."""
        for col_idx, header in enumerate(headers, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = len(str(header))
            for row in ws.iter_rows(
                min_row=2, max_row=ws.max_row,
                min_col=col_idx, max_col=col_idx
            ):
                for cell in row:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_len + 3, 50)

    def _write_summary_sheet(self):
        """
        Write a Summary sheet with counts, total processing time, and average time per document.
        Success and Failed counts are derived directly from Final Output (PASS / FAIL).
        """
        if not self._stats:
            return

        # If resuming, an old Summary sheet may already exist — remove it so
        # the new one (with correctly combined stats) replaces it cleanly
        # instead of being created as "Summary1" alongside a stale one.
        if "Summary" in self.wb.sheetnames:
            del self.wb["Summary"]
        ws = self.wb.create_sheet(title="Summary", index=0)  # insert at front

        # Clean Summary Headers (Test Pass / Test Fail merged into Success / Failed)
        summary_headers = [
            "Document Type", "Total Executed", "Success", "Failed", "Error",
            "Success Rate", "Total Time (s)", "Avg Time/Doc (s)"
        ]
        for col_idx, h in enumerate(summary_headers, start=1):
            self._write_header_cell(ws, 1, col_idx, h)
        ws.row_dimensions[1].height = 24

        grand_total = 0
        grand_success = 0
        grand_failed = 0
        grand_error = 0
        grand_time = 0.0

        current_row = 2
        # Data rows per document type
        for doc_type, counts in self._stats.items():
            total      = counts["TOTAL"]
            success    = counts["SUCCESS"]  # Final Output == PASS
            failed     = counts["FAILED"]   # Final Output == FAIL
            error      = counts["ERROR"]    # Final Output == REVIEW / ERROR
            total_time = counts.get("TOTAL_TIME", 0.0)
            avg_time   = (total_time / total) if total > 0 else 0.0
            rate_pct   = (success / total * 100) if total > 0 else 0.0
            rate_str   = f"{rate_pct:.1f}%"

            grand_total   += total
            grand_success += success
            grand_failed  += failed
            grand_error   += error
            grand_time    += total_time

            values = [
                doc_type,
                total,
                success,
                failed,
                error,
                rate_str,
                f"{total_time:.1f}s",
                f"{avg_time:.1f}s"
            ]
            for col_idx, val in enumerate(values, start=1):
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font      = CELL_FONT
                cell.alignment = CENTER_ALIGN if col_idx > 1 else LEFT_ALIGN
                cell.border    = THIN_BORDER
                # Colour by success rate
                if col_idx == 6:
                    rate_num = success / total if total > 0 else 0
                    cell.fill = (SUCCESS_FILL if rate_num >= 0.8
                                 else UNKNOWN_FILL if rate_num >= 0.5
                                 else FAILED_FILL)

            ws.row_dimensions[current_row].height = 20
            current_row += 1

        # ── Overall Summary Row at bottom ────────────────────
        grand_rate_pct = (grand_success / grand_total * 100) if grand_total > 0 else 0.0
        grand_avg_time = (grand_time / grand_total) if grand_total > 0 else 0.0

        total_row_values = [
            "OVERALL TOTAL",
            grand_total,
            grand_success,
            grand_failed,
            grand_error,
            f"{grand_rate_pct:.1f}%",
            f"{grand_time:.1f}s",
            f"{grand_avg_time:.1f}s"
        ]

        total_fill = PatternFill("solid", fgColor="D9E1F2")  # soft accent blue
        total_font = Font(name="Calibri", size=10, bold=True)
        for col_idx, val in enumerate(total_row_values, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font      = total_font
            cell.fill      = total_fill
            cell.alignment = CENTER_ALIGN if col_idx > 1 else LEFT_ALIGN
            cell.border    = THIN_BORDER

        ws.row_dimensions[current_row].height = 22

        # Column widths
        ws.column_dimensions["A"].width = 28
        for col in ["B", "C", "D", "E", "F", "G", "H"]:
            ws.column_dimensions[col].width = 18

        ws.freeze_panes = "A2"
        logger.info("Summary sheet written with overall timing statistics.")

    def _save(self):
        """Save the workbook (called after every row for crash safety)."""
        self.wb.save(str(self.filepath))
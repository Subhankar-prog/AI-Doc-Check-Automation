"""
create_input_map.py - Generate input_map.xlsx with one sheet per document type

Structure:
    Sheet "Aadhaar Card"                  <- paste all Aadhaar image paths here
    Sheet "PAN Card"                      <- paste all PAN image paths here
    Sheet "Indian Passport"               <- etc.
    Sheet "College Id"                    <- new
    Sheet "Destitute Certificate"         <- new
    Sheet "HIV Certificate"               <- new
    Sheet "Without Shelter Certificate"   <- new
    Sheet "Manual Scavanger Certificate" <- new
    Sheet "Vulnerable Tribal Certificate" <- new
    Sheet "Bonded Labour Certificate"     <- new
    Sheet "Single Mother Certificate"     <- new
    Sheet "Bank Passbook"                 <- new
    ...

Each sheet has:
    Column A: File Path  (you fill this in)
    Column B: Notes      (optional)

Run once:  python create_input_map.py
Then open input/input_map.xlsx and fill in your real paths.
"""
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

INPUT_DIR   = Path(__file__).parent / "input"
OUTPUT_FILE = INPUT_DIR / "input_map.xlsx"

# Styling
HEADER_FILL  = PatternFill("solid", fgColor="1E3A5F")
HEADER_FONT  = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
INSTR_FILL   = PatternFill("solid", fgColor="FFF3CC")
INSTR_FONT   = Font(italic=True, color="7F6000", name="Calibri", size=10)
DATA_FONT    = Font(name="Calibri", size=10)
SAMPLE_FONT  = Font(name="Calibri", size=10, italic=True, color="AAAAAA")
CENTER       = Alignment(horizontal="center", vertical="center")
LEFT         = Alignment(horizontal="left",   vertical="center")
THIN         = Side(style="thin")
THIN_BORDER  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# One distinct light colour per document type sheet tab
SHEET_COLORS = {
    # --- Original 10 types ---
    "Aadhaar Card":                  ("E8F4FD", "1565C0"),  # blue
    "PAN Card":                      ("FDE8F4", "880E4F"),  # pink
    "Indian Passport":               ("E8FDE8", "1B5E20"),  # green
    "Driving Licence":               ("FDF5E8", "E65100"),  # orange
    "10th Marksheet":                ("F0E8FD", "4A148C"),  # purple
    "12th Marksheet":                ("FDE8E8", "B71C1C"),  # red
    "Graduation Certificate":        ("E8FDF5", "004D40"),  # teal
    "Income Certificate":            ("FDEEE8", "BF360C"),  # deep orange
    "Caste Certificate":             ("E8EAFD", "1A237E"),  # indigo
    "Residence Certificate":         ("F5FDE8", "33691E"),  # light green

    # --- Newly added types (added 2026-09-05) ---
    "College Id":                    ("FFF8E1", "F57F17"),  # amber
    "Destitute Certificate":         ("FCE4EC", "880E4F"),  # deep pink
    "HIV Certificate":               ("F3E5F5", "6A1B9A"),  # deep purple
    "Without Shelter Certificate":   ("E0F7FA", "006064"),  # cyan
    "Manual Scavanger Certificate": ("FBE9E7", "BF360C"),  # deep orange 2
    "Vulnerable Tribal Certificate": ("E8F5E9", "2E7D32"),  # dark green
    "Bonded Labour Certificate":     ("EDE7F6", "4527A0"),  # deep purple 2
    "Single Mother Certificate":     ("FFF3E0", "E65100"),  # orange 2
    "Bank Passbook":                 ("E1F5FE", "01579B"),  # light blue
}

PLACEHOLDER_ROWS = 10   # blank rows pre-created when folder is empty

# Maps each document type -> its Test Data subfolder name.
# The folder is scanned at creation time; found files are pre-filled
# into Column A so you don't have to type paths manually.
TEST_DATA_DIR = Path(__file__).parent / "Test Data"
TEST_DATA_FOLDERS = {
    # --- Original 10 types ---
    "Aadhaar Card":                  "Adhar card  for Testing",
    "PAN Card":                      "PAN card for Testing",
    "Indian Passport":               "Passport for Testing",
    "Driving Licence":               "DL for Testing",
    "10th Marksheet":                "10th Marksheet for Testing",
    "12th Marksheet":                "12th Certificate for Testing",
    "Graduation Certificate":        "Degree Certificate",
    "Income Certificate":            "Income Certificate for Testing",
    "Caste Certificate":             "Cast Certificate for Testing",
    "Residence Certificate":         "Residence Certificate for Testing",

    # --- Newly added types ---
    "College Id":                    "College Id for Testing",
    "HIV Certificate":               "HIV Certificate for Testing",
    "Vulnerable Tribal Certificate": "Vulnerable Tribal Certificate for Testing",
    "Bonded Labour Certificate":     "Bonded Labour Certificate for Testing",
    "Single Mother Certificate":     "Single Mother Certificate for Testing",
    "Bank Passbook":                 "Bank Passbook for Testing",
    "Manual Scavanger Certificate":  "Manual Scavengers Certificate for Testing",
    "Without Shelter Certificate":   "Without Shelter certificate for Testing",
    "Destitute Certificate":         "Destitute Certificate for Testing",
}


def _scan_folder(doc_type: str):
    """
    Return sorted list of absolute file paths found in the Test Data folder
    for this document type. Only picks up allowed extensions.
    Returns [] if folder is missing or empty.
    """
    ALLOWED = {".jpg", ".jpeg", ".png", ".webp", ".pdf"}
    folder_name = TEST_DATA_FOLDERS.get(doc_type)
    if not folder_name:
        return []
    folder = TEST_DATA_DIR / folder_name
    if not folder.is_dir():
        return []
    files = sorted(
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in ALLOWED
    )
    return [str(p.resolve()) for p in files]


def _write_header(ws, doc_type: str, bg: str, accent: str):
    """Write top instruction row + column headers for a doc-type sheet."""
    # Row 1: instruction banner (spans all 4 cols)
    ws.merge_cells("A1:D1")
    cell = ws["A1"]
    cell.value = (
        f"Paste full file paths in Column A (JPG / PNG / PDF). "
        f"Column B = optional notes. "
        f"Column C = Run Status (auto-updated by script). "
        f"All files in this sheet = {ws.title}"
    )
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.font      = Font(italic=True, color=accent, name="Calibri", size=10)
    cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 36

    # Row 2: column headers
    headers = ["File Path", "Notes (optional)", "Run Status", "Output Row"]
    header_bg = PatternFill("solid", fgColor="1E3A5F")
    header_ft = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=2, column=col_idx, value=h)
        c.fill      = header_bg
        c.font      = header_ft
        c.alignment = CENTER
        c.border    = THIN_BORDER
    ws.row_dimensions[2].height = 20

    # Freeze below header
    ws.freeze_panes = "A3"


def _write_placeholder_rows(ws, doc_type: str, bg: str, n: int = PLACEHOLDER_ROWS):
    """Pre-fill n blank rows with styled placeholder text and PENDING status."""
    bg_fill      = PatternFill("solid", fgColor=bg)
    pending_fill = PatternFill("solid", fgColor="F5F5F5")
    pending_font = Font(name="Calibri", size=10, italic=True, color="AAAAAA")
    status_font  = Font(name="Calibri", size=10, italic=True, color="888888")
    center_align = Alignment(horizontal="center", vertical="center")

    for i in range(n):
        row = i + 3   # data starts at row 3

        # Col A: file path placeholder
        a = ws.cell(row=row, column=1,
                    value=f"<path to {doc_type} file {i+1}>")
        a.fill      = bg_fill
        a.font      = SAMPLE_FONT
        a.alignment = LEFT
        a.border    = THIN_BORDER

        # Col B: notes
        b = ws.cell(row=row, column=2, value="")
        b.fill      = bg_fill
        b.font      = DATA_FONT
        b.alignment = LEFT
        b.border    = THIN_BORDER

        # Col C: Run Status = PENDING
        c = ws.cell(row=row, column=3, value="PENDING")
        c.fill      = pending_fill
        c.font      = status_font
        c.alignment = center_align
        c.border    = THIN_BORDER

        # Col D: Output Row = blank
        d = ws.cell(row=row, column=4, value="")
        d.fill      = bg_fill
        d.font      = DATA_FONT
        d.alignment = center_align
        d.border    = THIN_BORDER

        ws.row_dimensions[row].height = 18


def create():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # remove default empty sheet

    total_files = 0
    for doc_type, (bg, accent) in SHEET_COLORS.items():
        ws = wb.create_sheet(title=doc_type)

        # Sheet tab colour (visible in Excel)
        ws.sheet_properties.tabColor = bg

        # Column widths
        ws.column_dimensions["A"].width = 65
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 12

        _write_header(ws, doc_type, bg, accent)

        # --- Auto-fill from Test Data folder ---
        files = _scan_folder(doc_type)
        if files:
            bg_fill  = PatternFill("solid", fgColor=bg)
            for i, fpath in enumerate(files):
                row = i + 3  # data starts at row 3
                a = ws.cell(row=row, column=1, value=fpath)
                a.fill      = bg_fill
                a.font      = DATA_FONT
                a.alignment = LEFT
                a.border    = THIN_BORDER
                for col in (2, 3, 4):
                    c = ws.cell(row=row, column=col,
                                value="PENDING" if col == 3 else "")
                    c.fill      = PatternFill("solid", fgColor="F5F5F5") if col == 3 else bg_fill
                    c.font      = Font(name="Calibri", size=10,
                                      italic=True, color="888888") if col == 3 else DATA_FONT
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.border    = THIN_BORDER
                ws.row_dimensions[row].height = 18
            total_files += len(files)
            print(f"  [{doc_type}]  {len(files)} file(s) pre-filled from Test Data folder")
        else:
            # Folder empty or not yet populated — use placeholder rows
            _write_placeholder_rows(ws, doc_type, bg)
            print(f"  [{doc_type}]  no files found — placeholder rows added")

    wb.save(OUTPUT_FILE)

    print()
    print(f"[OK] input_map.xlsx created at: {OUTPUT_FILE}")
    print(f"     {total_files} file path(s) pre-filled from Test Data folders.")
    print()
    print("HOW TO FILL (for empty sheets):")
    print("  1. Open the sheet for your document type (e.g. 'Aadhaar Card')")
    print("  2. Replace the grey placeholder paths in Column A with real file paths")
    print("  3. You can add as many rows as needed (100s is fine)")
    print("  4. Leave sheets EMPTY if you have no files for that type")
    print("  5. Save, then run:  python main.py")


if __name__ == "__main__":
    create()

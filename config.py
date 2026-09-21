# ============================================================
#  config.py  -  Central configuration for AI Doc Automation
#  Portal: https://aiprojects.odisha.gov.in/demo  (VERIFAI)
# ============================================================

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# --- Paths -----------------------------------------------------
# Each can be overridden in .env. If not set there, these defaults are used
# (same folders as before — nothing changes unless you add them to .env).
BASE_DIR        = Path(__file__).parent.resolve()
INPUT_MAP_FILE  = BASE_DIR / "input" / "input_map.xlsx"

OUTPUT_DIR      = Path(os.getenv("OUTPUT_DIR", str(BASE_DIR / "output")))
LOGS_DIR        = Path(os.getenv("LOGS_DIR", str(BASE_DIR / "logs")))
TEST_DATA_DIR   = Path(os.getenv("TEST_DATA_DIR", str(BASE_DIR / "Test Data")))

SCREENSHOTS_DIR = OUTPUT_DIR / "screenshots"

# --- Portal --------------------------------------------------
PORTAL_URL = os.getenv("PORTAL_URL", "https://aiprojects.odisha.gov.in/demo")

# --- Timeouts ------------------------------------------------
PAGE_LOAD_TIMEOUT  = 30       # seconds for page load
ENGINE_TIMEOUT     = 150      # seconds to wait for engine result (max)
POLL_INTERVAL      = 2        # seconds between status polls
RETRY_ON_TIMEOUT   = 0        # number of retries if engine times out
AUTO_RETRY_FAILED_PASSES = 1  # retry passes for FAILED/ERROR/TIMEOUT rows after main run
                              # (0 = disabled; override with --retry-passes N on CLI)
SKIP_DONE = True              # True  = skip SUCCESS/SKIPPED rows by default (recommended)
                              # False = re-run every row from scratch every time
                              # Note: --resume also skips done AND merges into previous output Excel

# --- Browser -------------------------------------------------
HEADLESS_MODE = False        # Set True to hide Chrome window

# --- Portal's stated upload rules (used for a PRE-upload prediction only —
#     the file is still always uploaded so the real portal behavior is
#     verified against this prediction, never skipped) -------------------
MAX_UPLOAD_SIZE_MB       = 5
ALLOWED_UPLOAD_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".pdf"}

# --- Document Type Dropdown Values ---------------------------
# Keys   = value you write in input_map.xlsx "Document Type" column (also used for the sheet tab name)
# Values = exact text shown in the portal's "Document Type" dropdown menu
# 
# 👉 TO ADD A NEW TYPE TO THE DROPDOWN:
# Simply add a new line here. 
# Format: "Your Excel Sheet Name": "Exact Text in Portal Dropdown",
DOCUMENT_TYPES = {
    # --- Original 10 types ---
    "Aadhaar Card":                  "Aadhaar Card",
    "PAN Card":                      "PAN Card",
    "Indian Passport":               "Indian Passport",
    "Driving Licence":               "Driving Licence",
    "10th Marksheet":                "10th Marksheet",
    "12th Marksheet":                "12th Marksheet",
    "Graduation Certificate":        "Graduation Certificate",
    "Income Certificate":            "Income Certificate",
    "Caste Certificate":             "Caste Certificate",
    "Residence Certificate":         "Residence Certificate",

    # --- Newly added types (added 2026-09-05) ---
    "College Id":                    "College Id",
    "Destitute Certificate":         "Destitute Certificate",
    "HIV Certificate":               "HIV Certificate",
    "Without Shelter Certificate":   "Without Shelter Certificate",
    "Manual Scavanger Certificate": "Manual Scavenger Certificate",
    "Vulnerable Tribal Certificate": "Vulnerable Tribal Certificate",
    "Bonded Labour Certificate":     "Bonded Labour Certificate",
    "BOC Certificate":               "BOC Certificate",
    "Single Mother Certificate":     "Single Mother Certificate",
    "Bank Passbook":                 "Bank Passbook",

    # --- Newly added types (added 2026-09-16) ---
    "General Caste Self Decl":        "General Caste Self Declaration",
    "Institute Declaration Certificate":     "Institute Declaration Certificate",
    "Orphan Certificate":             "Orphan Certificate",
    "Ration Card":                    "Ration Card",

     # --- Newly added types (added 2026-09-19) ---
    "Divyang Certificate":            "Divyang Certificate",
    "Occupation Certificate":         "Occupation Certificate",
    "Annual Contribution Certificate":"Annual Contribution Certificate",
    "Transgender Certificate":        "Transgender Certificate",
    "Visa":                           "Visa",
    "Parent Declaration Certificate": "Parent Declaration Certificate"
}

# --- XPaths (verified from VERIFAI portal console) -----------
XPATHS = {
    # --- Sidebar Navigation ---
    "verify_document_link": (
        "//div[contains(@class, 'shell')]//*[contains(text(), 'Verify Document')]"
    ),
    "start_verification": (
        "//div[contains(@class, 'hero')]//*[contains(text(), 'Start verification')]"
    ),

    # --- Step 01: Document Type <select> ---
    "doc_type_select":          "(//select)[1]",
    "doc_type_select_labelled": (
        "//*[contains(text(),'DOCUMENT TYPE') or contains(text(),'Document Type')]"
        "/following::select[1]"
    ),

    # --- Step 02: File Upload ---
    # The portal uses a hidden <input type="file">; we make it visible via JS.
    "file_input": "//input[@type='file']",

    # --- Step 03: Plain Extraction (custom <button class="mode">) ---
    # Confirmed from console: shell > Plain Extraction text element
    "plain_extraction_btn": (
        "//div[contains(@class, 'shell')]//*[contains(text(), 'Plain Extraction')]"
    ),
    "plain_extraction_btn_v2": (
        "//button[contains(@class,'mode')][.//*[contains(text(),'Plain extraction')] "
        "or contains(text(),'Plain extraction')]"
    ),

    # --- Submit to Engine ---
    # Confirmed from console: submit-row class > Submit to engine text
    "submit_to_engine": (
        "//div[contains(@class, 'submit-row')]//*[contains(text(), 'Submit to engine')]"
    ),
    "submit_to_engine_v2": "//*[contains(text(), 'Submit to engine')]",

    # --- Processing indicator (only on the submit button itself) ---
    "processing_indicator": (
        "//button[contains(@class,'submit') or contains(@class,'btn') or @type='submit']"
        "[contains(text(), 'Processing') or contains(., 'Processing')]"
    ),

    # --- Result Page Tabs ---
    # Confirmed from console: result-tabs class > Overview / Extraction text
    "overview_tab": (
        "//div[contains(@class, 'result-tabs')]//*[contains(text(), 'Overview')]"
    ),
    "extraction_tab": (
        "//div[contains(@class, 'result-tabs')]//*[contains(text(), 'Extraction')]"
    ),

    # --- Result page: wait for "TRANSACTION COMPLETE" or tab area or New submission button ---
    "result_appeared": (
        "//*[contains(text(),'TRANSACTION COMPLETE') or "
        "contains(text(),'Pipeline failure') or "
        "contains(text(),'Plain extraction result') or "
        "contains(@class,'result-tabs') or "
        "contains(text(),'New submission')]"
    ),

    # --- Overall verdict on Overview tab ---
    # Screenshot shows prominent "SUCCESS" text inside a verdict box
    "verdict_box": (
        "//*[contains(@class,'verdict') or contains(@class,'overall')]"
    ),

    # --- Extraction tab field rows ---
    # Screenshot shows rows: label (left, grey text) | value (right, bold dark text)
    # These rows are inside the Extractable Fields section.
    # Strategy A - table rows
    "extract_rows_table": "//table//tr[td]",
    # Strategy B - generic row divs (most likely for this SPA)
    "extract_container": (
        "//*[contains(@class,'extractable') or contains(@class,'fields') "
        "or contains(@class,'extraction')]"
    ),
    # Strategy C - definition list
    "extract_rows_dl": "//dl/dt",

    # --- New submission button ---
    # Confirmed from console: result-actions class > New submission text
    "new_submission": (
        "//div[contains(@class, 'result-actions')]//*[contains(text(), 'New submission')]"
    ),
    "new_submission_v2": (
        "//*[contains(text(), 'New submission') or contains(text(), 'New Submission')]"
    ),
}

# --- Excel Output --------------------------------------------
OUTPUT_EXCEL_PREFIX = "ai_doc_results"

# Column names used in input_map.xlsx
INPUT_EXCEL_COLUMNS = {
    "file_path": "File Path",      # Column A: absolute path to image/pdf
    "doc_type":  "Document Type",  # Column B: document type (must match DOCUMENT_TYPES key)
}

# These columns always appear first in every sheet tab
FIXED_COLUMNS = [
    "Filename",
    "Document Type",
    "Expected Status",
    "Actual Status",
    "Final Output",
    "Error Message",
]

# Error message substrings that indicate the portal correctly rejected bad
# input (i.e. an expected/intentional negative-test scenario), rather than
# a genuine bug. Matched case-insensitively against the Error Message text.
KNOWN_VALIDATION_REJECTION_PATTERNS = [
    "multiple_pages",
    "multiple pages",
    "document_type_mismatch",
    "document type mismatch",
    "upload a jpeg, png, webp, or pdf",
    "up to 9 mb",
    "up to 5 mb",
]
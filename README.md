# VERIFAI Document Extraction Automation

Automates bulk document extraction on [https://aiprojects.odisha.gov.in/demo](https://aiprojects.odisha.gov.in/demo) using **Selenium + Python**.
The script processes documents one by one, extracts all fields, and saves results to a formatted Excel file automatically.

---

## Project Structure

```
ai_doc_automation/
│
├── main.py                 ← Run this to start the automation
├── config.py               ← All XPaths, URLs, timeouts (edit here if portal changes)
├── driver_setup.py         ← Chrome WebDriver setup
├── doc_uploader.py         ← Upload, select type, submit, wait for result
├── result_parser.py        ← Scrape extracted fields from result page
├── excel_writer.py         ← Write results to output Excel
├── input_tracker.py        ← Track progress live in input Excel (Run Status column)
├── create_input_map.py     ← Run once to create the input template
│
├── input/
│   └── input_map.xlsx      ← YOUR INPUT: one sheet per document type
│
├── output/
│   └── ai_doc_results_<timestamp>.xlsx   ← AUTO-CREATED: all extracted results
│   └── screenshots/                      ← Auto-saved screenshot per document
│
└── logs/
    └── run_<timestamp>.log               ← Full run log
```

---

## Quick Start

### Step 1 — Install dependencies
```bash
pip install -r requirements.txt
```

### Step 2 — Create the input Excel template
```bash
python create_input_map.py
```
This creates `input/input_map.xlsx` with **10 sheet tabs** — one per document type.

### Step 3 — Fill in your file paths
Open `input/input_map.xlsx`.
Each sheet tab corresponds to a document type:

| Sheet Tab | What to do |
|---|---|
| **Aadhaar Card** | Paste paths to all Aadhaar images in Column A |
| **PAN Card** | Paste paths to all PAN images in Column A |
| **Graduation Certificate** | Paste paths to all Graduation cert images in Column A |
| *(and so on for all 10 types)* | Leave any sheet empty if you have no files for it |

**Column layout in each sheet:**

| A — File Path | B — Notes (optional) | C — Run Status | D — Output Row |
|---|---|---|---|
| `D:\docs\aadhaar\file1.jpg` | | PENDING | |
| `D:\docs\aadhaar\file2.pdf` | | PENDING | |

> Column C and D are managed automatically by the script. Do not edit them.

### Step 4 — Run

```bash
# Run all document types
python main.py

# Run only one document type
python main.py --type "Graduation Certificate"
python main.py --type "Aadhaar Card"
python main.py --type "PAN Card"

# Check file paths first without opening browser
python main.py --dry-run

# Resume after stopping — continue where you left off
python main.py --resume

# Resume, but only for one specific document type

python main.py --resume --type "Aadhaar Card"

# Resume into a specific output file (instead of auto-detected latest)

python main.py --resume --resume-output "outputs/ai_doc_results_20260822_100850.xlsx"

# Run without visible Chrome window
python main.py --headless

# Resume + headless together — common for long unattended overnight runs
python main.py --resume --headless

# Use a different input file
python main.py --input "D:\my_docs\batch2.xlsx"

# Resume, one doc type, custom input, headless
python main.py --resume --type "PAN Card" --input "batch2_map.xlsx" --headless
```

---

## How It Works (Step by Step)

```
For each file in input_map.xlsx:

  1.  Opens Chrome → navigates to portal → clicks "Verify Document"
  2.  Selects the document type from the dropdown
  3.  Clicks "Plain extraction" radio button
  4.  Uploads the document file (no popup dialogs — done via code)
  5.  Verifies dropdown and radio are correct before submitting
  6.  Clicks "Submit to engine"
  7.  Waits up to 3 minutes, polling every 5 seconds for result
  8.  Reads "OVERALL VERDICT" from the Overview tab
  9.  If SUCCESS → clicks Extraction tab → scrapes all label-value fields
  10. Writes one row to the output Excel immediately (crash-safe)
  11. Updates "Run Status" in input Excel (SUCCESS / FAILED / ERROR)
  12. Takes a screenshot of the result page
  13. Clicks "New submission" → resets form → moves to next file
```

---

## Input Excel (`input/input_map.xlsx`)

- **10 sheet tabs** — one per document type
- **Column A** = full file path to image or PDF (`JPG / PNG / PDF`)
- **Column B** = optional notes (you can leave blank)
- **Column C** = Run Status (written by script — do not edit)
- **Column D** = Output row reference (written by script)

You can add **as many rows as needed** per sheet — 100s of files per type is fine.
Leave any sheet completely empty if you have no files for that document type.

**Valid document types (sheet names):**
- Aadhaar Card
- PAN Card
- Indian Passport
- Driving Licence
- 10th Marksheet
- 12th Marksheet
- Graduation Certificate
- Income Certificate
- Caste Certificate
- Residence Certificate

---

## Run Status Tracking (Column C in input Excel)

As the script runs, Column C is updated live for every row:

| Status | Colour | Meaning |
|---|---|---|
| `PENDING` | Grey | Not yet processed |
| `RUNNING` | Yellow | Currently uploading and processing |
| `SUCCESS` | Green | Done — fields extracted successfully |
| `FAILED` | Red | Engine returned a failure for this file |
| `ERROR` | Orange | Script exception occurred |
| `TIMEOUT` | Purple | Engine did not respond within 3 minutes |

If the script crashes mid-run, open `input_map.xlsx` to see exactly which files completed and which are still pending. Then run with `--resume` to continue without repeating completed files.

---

## Output Excel (`output/ai_doc_results_<timestamp>.xlsx`)

The output Excel is **created automatically** when the script runs. You do not need to create it.

**Structure:**

```
ai_doc_results_20260820_153000.xlsx
│
├── Sheet: Summary                    ← totals per document type
├── Sheet: Aadhaar Card               ← all Aadhaar results
├── Sheet: PAN Card                   ← all PAN results
├── Sheet: Graduation Certificate     ← all Graduation results
└── ...                               ← one sheet per type processed
```

**Each sheet contains:**

| Filename | Status | *...all extracted fields...* | Error Message | Screenshot | Processing Time | Timestamp |
|---|---|---|---|---|---|---|
| grad_001.jpg | SUCCESS | Grade: A+, Result: FIRST CLASS..., Student Name: ... | — | path/to/ss.png | 47.3s | 2026-08-20 10:00 |
| grad_002.jpg | FAILED | — | Image too blurry | path/to/ss.png | 12.1s | 2026-08-20 10:02 |

- Extracted field columns are **created dynamically** — whatever the portal returns becomes a column
- Different document types have different field columns
- 🟢 **Green rows** = SUCCESS
- 🔴 **Red rows** = FAILED / ERROR / TIMEOUT
- 🟡 **Yellow rows** = UNKNOWN
- A new timestamped file is created each run — previous results are never overwritten

---

## Updating Selectors

If the portal UI changes, update `config.py`:

```python
XPATHS = {
    "doc_type_select":      "...",   # document type <select> dropdown
    "plain_extraction_btn": "...",   # Plain extraction button
    "file_input":           "...",   # <input type='file'>
    "submit_to_engine":     "...",   # Submit to engine button
    "extraction_tab":       "...",   # Extraction tab
    "new_submission":       "...",   # New submission button
}
```

**How to find XPaths in Chrome:**
1. Press **F12** → click the cursor/inspector icon
2. Click the element on the portal page
3. In Elements panel → right-click the highlighted tag → **Copy → Copy XPath**
4. Paste into `config.py`

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Chrome doesn't open | Make sure Google Chrome is installed |
| `webdriver-manager` error | Run `pip install --upgrade webdriver-manager` |
| File not uploading | Check Column A has the full absolute path (not relative) |
| Dropdown not selecting | Inspect the `<select>` in DevTools, update `doc_type_select` XPath in `config.py` |
| Engine times out | Increase `ENGINE_TIMEOUT` in `config.py` (default is 180 seconds) |
| Fields not extracted | Check logs — the parser tries 5 strategies; log shows which one matched |
| Script crashes mid-run | Output Excel is saved after every row — partial results are safe. Run with `--resume` to continue |
| input_map.xlsx locked | Close the file in Excel before running the script, or the Run Status column cannot be written |

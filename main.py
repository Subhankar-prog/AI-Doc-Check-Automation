"""
main.py - Orchestrator for AI Document Check Automation (VERIFAI)
Portal: https://aiprojects.odisha.gov.in/demo

Usage:
    python main.py                               # Run all document types
    python main.py --type "Graduation Certificate"  # Run one type only
    python main.py --resume                      # Skip already-done rows
    python main.py --dry-run                     # Validate paths only
    python main.py --headless                    # Hide browser window
    python main.py --input my_map.xlsx           # Custom input file
"""
import sys
import time
import logging
import argparse
from pathlib import Path

import openpyxl
from tqdm import tqdm

from config import (
    PORTAL_URL, INPUT_MAP_FILE, INPUT_EXCEL_COLUMNS, DOCUMENT_TYPES,
    LOGS_DIR, RETRY_ON_TIMEOUT
)
from driver_setup   import create_driver
from doc_uploader   import (
    navigate_to_upload, select_document_type,
    ensure_plain_extraction, upload_file,
    verify_and_submit, wait_for_result,
    click_new_submission, take_screenshot,
    UploadRejectedError
)
from result_parser  import parse_result
from excel_writer   import ExcelWriter
from input_tracker  import InputTracker

logger = None   # set in main()


# ──────────────────────────────────────────────────────────
# Logging Setup
# ──────────────────────────────────────────────────────────
def setup_logging():
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    ts       = time.strftime("%Y%m%d_%H%M%S")
    log_file = LOGS_DIR / f"run_{ts}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  [%(levelname)-8s]  %(name)s - %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ]
    )
    return logging.getLogger("main")


# ──────────────────────────────────────────────────────────
# Read Input Excel (delegates to InputTracker)
# ──────────────────────────────────────────────────────────
def load_records(input_file: Path, doc_type_filter: str, skip_done: bool) -> list:
    """Load records from input Excel via InputTracker."""
    tracker = InputTracker(input_file)
    records = tracker.load_records(
        doc_type_filter=doc_type_filter,
        skip_done=skip_done
    )
    return tracker, records


# ──────────────────────────────────────────────────────────
# Process a single document (with retry on timeout)
# ──────────────────────────────────────────────────────────
def process_document(driver, record: dict, excel: ExcelWriter,
                     tracker: InputTracker, attempt: int = 1):
    file_path  = record["file_path"]
    doc_type   = record["doc_type"]
    filename   = Path(file_path).name
    sheet_name = record["sheet_name"]
    row_num    = record["row_num"]

    logger.info("=" * 60)
    logger.info(f"Processing [{record.get('_idx','?')}/{record.get('_total','?')}]: "
                f"{filename}  |  {doc_type}  |  Attempt {attempt}")
    logger.info("=" * 60)

    # Mark as RUNNING in input Excel immediately
    tracker.mark_running(sheet_name, row_num)

    start_time = time.time()

    try:
        # Step 1: Select document type
        select_document_type(driver, doc_type)

        # Step 2: Ensure Plain Extraction is selected
        ensure_plain_extraction(driver)

        # Step 3: Upload file (raises UploadRejectedError if size/format error banner appears)
        upload_file(driver, file_path)

        # Step 4: Verify state and submit
        verify_and_submit(driver, doc_type)

        # Step 5: Wait for result (detected by TRANSACTION COMPLETE / Pipeline failure / result-tabs)
        result_appeared = wait_for_result(driver)
        elapsed = time.time() - start_time

        if not result_appeared:
            if attempt <= RETRY_ON_TIMEOUT:
                logger.warning(f"Timeout on attempt {attempt}. Retrying...")
                tracker.mark_running(sheet_name, row_num)   # still running
                click_new_submission(driver)
                return process_document(driver, record, excel, tracker, attempt + 1)
            else:
                take_screenshot(driver, f"timeout_{filename}")
                result = {
                    "Filename":         filename,
                    "Document Type":    doc_type,
                    "Status":           "TIMEOUT",
                    "Error Message":    f"Engine did not respond after {elapsed:.0f}s",
                    "_processing_time": round(elapsed, 1),
                }
        else:
            # Step 6: Parse result page (accurately detects SUCCESS vs FAILED)
            result = parse_result(driver, filename, doc_type, elapsed)

    except UploadRejectedError as ure:
        elapsed = time.time() - start_time
        logger.warning(f"File upload rejected for {filename}: {ure}")
        try:
            take_screenshot(driver, f"upload_rejected_{filename}")
        except Exception:
            pass
        result = {
            "Filename":         filename,
            "Document Type":    doc_type,
            "Status":           "FAILED",
            "Error Message":    f"Upload rejected: {ure}",
            "_processing_time": round(elapsed, 1),
        }

    except Exception as e:
        elapsed = time.time() - start_time
        logger.error(f"Exception on {filename}: {e}", exc_info=True)
        try:
            take_screenshot(driver, f"error_{filename}")
        except Exception:
            pass
        result = {
            "Filename":         filename,
            "Document Type":    doc_type,
            "Status":           "ERROR",
            "Error Message":    str(e),
            "_processing_time": round(elapsed, 1),
        }


    # Step 7: Write to OUTPUT Excel (crash-safe, saves immediately)
    excel.append_row(result)

    # Step 8: Update INPUT Excel status column
    status = result.get("Status", "ERROR")
    tracker.mark_done(sheet_name, row_num, status)

    icon = "[OK]" if status == "SUCCESS" else "[FAIL]"
    logger.info(f"{icon} {filename}: {status} ({result['Processing Time (s)']}s)")

    return result



# ──────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────
def main():
    global logger

    parser = argparse.ArgumentParser(
        description="VERIFAI Document Automation",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument(
        "--input", type=str, default=str(INPUT_MAP_FILE),
        help="Path to input Excel map file"
    )
    parser.add_argument(
        "--type", type=str, default=None, dest="doc_type",
        help=(
            "Run only one document type (sheet name).\n"
            'Examples:\n'
            '  python main.py --type "Graduation Certificate"\n'
            '  python main.py --type "Aadhaar Card"'
        )
    )
    parser.add_argument(
        "--resume", action="store_true",
        help=(
            "Skip rows already marked SUCCESS/FAILED in input Excel.\n"
            "Use this to continue after a crash without re-processing done files."
        )
    )
    parser.add_argument(
        "--headless", action="store_true",
        help="Run Chrome without visible window"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Check file paths only — no browser, no uploads"
    )
    args = parser.parse_args()

    logger = setup_logging()
    logger.info("=" * 60)
    logger.info("  VERIFAI Document Automation - Starting")
    logger.info("=" * 60)

    input_file = Path(args.input)
    if not input_file.exists():
        logger.error(f"Input file not found: {input_file}")
        logger.error("Run:  python create_input_map.py  to create the template.")
        sys.exit(1)

    # ── Load records via tracker ───────────────────────────
    logger.info(f"Input file : {input_file}")
    if args.resume:
        logger.info("[--resume] Skipping rows already marked SUCCESS/FAILED.")
    if args.doc_type:
        logger.info(f"[--type]   Running only: '{args.doc_type}'")

    tracker, records = load_records(input_file, args.doc_type, args.resume)
    total = len(records)

    if total == 0:
        logger.error(
            "No pending records found.\n"
            "  - Check that you filled file paths in input_map.xlsx\n"
            "  - If using --resume, all rows may already be marked done"
        )
        sys.exit(1)

    logger.info(f"Records to process: {total}")

    # ── Dry run ───────────────────────────────────────────
    if args.dry_run:
        logger.info("[DRY RUN] Checking file paths...")
        all_ok  = True
        missing = 0
        for r in records:
            exists = Path(r["file_path"]).exists()
            tag    = "[OK]    " if exists else "[MISSING]"
            logger.info(f"  {tag} {r['doc_type']:30s}  {r['file_path']}")
            if not exists:
                all_ok = False
                missing += 1

        logger.info(
            f"\n[DRY RUN] Done.  "
            f"Found: {total - missing}  |  Missing: {missing}"
        )
        sys.exit(0 if all_ok else 1)

    # ── Headless mode ─────────────────────────────────────
    if args.headless:
        import config
        config.HEADLESS_MODE = True

    # ── Start browser + output Excel ──────────────────────
    driver = create_driver()
    excel  = ExcelWriter()
    stats  = {"success": 0, "failed": 0, "error": 0}

    try:
        logger.info(f"Opening portal: {PORTAL_URL}")
        driver.get(PORTAL_URL)
        time.sleep(3)
        take_screenshot(driver, "01_dashboard")

        navigate_to_upload(driver)
        take_screenshot(driver, "02_verify_form")

        # ── Main loop ─────────────────────────────────────
        for idx, record in enumerate(tqdm(records, desc="Documents", unit="doc"), start=1):
            record["_idx"]   = idx
            record["_total"] = total

            logger.info(f"\n[{idx}/{total}] {record['file_path']}")

            try:
                result = process_document(driver, record, excel, tracker)
            except Exception as e:
                logger.error(f"Top-level exception processing document: {e}")
                result = {"Status": "ERROR", "Processing Time (s)": 0}

            s = result.get("Status", "ERROR").upper()
            if s == "SUCCESS":
                stats["success"] += 1
            elif s in ("FAILED", "TIMEOUT"):
                stats["failed"] += 1
            else:
                stats["error"] += 1

            # Reset form for next file (skip on last)
            if idx < total:
                try:
                    click_new_submission(driver)
                    time.sleep(1)
                except Exception as ex:
                    logger.warning(f"Error resetting form ({ex}). Restarting browser session...")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    driver = create_driver()
                    driver.get(PORTAL_URL)
                    time.sleep(2)
                    navigate_to_upload(driver)


        # ── Finish ────────────────────────────────────────
        output_path = excel.finalize()
        logger.info("\n" + "=" * 60)
        logger.info("  COMPLETE")
        logger.info(f"  Total   : {total}")
        logger.info(f"  Success : {stats['success']}")
        logger.info(f"  Failed  : {stats['failed']}")
        logger.info(f"  Errors  : {stats['error']}")
        logger.info(f"  Output  : {output_path}")
        logger.info("=" * 60)

    except KeyboardInterrupt:
        logger.warning("Interrupted by user (Ctrl+C). Saving partial results...")
        output_path = excel.finalize()
        logger.info(f"Partial results saved: {output_path}")
        logger.info("Run with --resume to continue from where you stopped.")

    finally:
        driver.quit()
        logger.info("Browser closed.")


if __name__ == "__main__":
    main()

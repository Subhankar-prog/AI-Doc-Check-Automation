"""
result_parser.py - Scrape results from VERIFAI result page

From the portal screenshots:
Overview tab: shows OVERALL VERDICT = SUCCESS/FAILED
Extraction tab: shows a list of "EXTRACTABLE FIELDS" as label | value rows
  - Label: left cell, grey text (e.g. "Grade", "Result", "Student Name")
  - Value: right cell, bold/colored (e.g. "A+", "FIRST CLASS WITH DISTINCTION")
  - Some values are "Not detected"

FAST-PATH LOGIC:
The Extraction tab only exists on the result page when the document was
processed successfully. So instead of always reading the Overview tab's
verdict text (a tab click + sleep + several XPath/body-text scans) and
THEN deciding whether to open Extraction, we try to open Extraction first:

  - Extraction tab present  -> it's a SUCCESS, scrape fields directly.
    No Overview click, no verdict-text reading, no body-text scan.
  - Extraction tab absent   -> genuinely failed/unknown, NOW it's worth
    paying for the slower Overview / verdict / failure-reason read.

This means the expensive diagnostic path only runs for the minority of
documents that actually failed, instead of running for every document.
"""

import time
import logging
from typing import Dict, Tuple

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

from config import XPATHS
from doc_uploader import _click_safe, take_screenshot, _find_element

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────
# Detect overall verdict from Overview tab (slow path — only used
# when the Extraction tab is not present, i.e. failure/unknown cases)
# ─────────────────────────────────────────────────────────────
def get_status(driver: webdriver.Chrome) -> Tuple[str, str]:
    """
    Click the Overview tab and read the overall verdict.

    Returns:
        (status, error_message)
        status: 'SUCCESS' | 'FAILED' | 'UNKNOWN'
    """
    logger.info("Reading verdict from Overview tab...")

    # First check for prominent pipeline failure / error banners on page
    err_banner = _get_pipeline_error(driver)
    if err_banner:
        logger.info(f"Verdict: FAILED (Pipeline Failure banner detected: {err_banner})")
        return "FAILED", err_banner

    # Click Overview tab if present
    try:
        overview = _find_element(driver, XPATHS["overview_tab"], timeout=5)
        _click_safe(driver, overview)
        time.sleep(0.4)
    except NoSuchElementException:
        logger.warning("Overview tab not found — reading current page.")

    # 1. Check for specific OVERALL VERDICT card text
    try:
        verdict_elements = driver.find_elements(
            By.XPATH,
            "//*[contains(text(),'OVERALL VERDICT') or contains(text(),'Overall Verdict') "
            "or contains(@class,'verdict')]/ancestor::*[contains(@class,'card') or contains(@class,'box') or contains(@class,'result') or position() <= 2][1]"
        )
        for vel in verdict_elements:
            vtext = vel.text.upper()
            if "FAILED" in vtext or "FAILURE" in vtext:
                err = _get_error_text(driver)
                logger.info(f"Verdict: FAILED (from verdict card) - {err}")
                return "FAILED", err
            elif "SUCCESS" in vtext:
                logger.info("Verdict: SUCCESS (from verdict card)")
                return "SUCCESS", ""
    except Exception as e:
        logger.debug(f"Verdict card check: {e}")

    # 2. Check full page text carefully
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text.upper()
    except Exception:
        body_text = ""

    if "PIPELINE FAILURE" in body_text or "PIPELINE FAILED" in body_text:
        err = _get_error_text(driver)
        logger.info(f"Verdict: FAILED (Pipeline failure text) - {err}")
        return "FAILED", err

    if "OVERALL VERDICT" in body_text:
        # Check which word appears closest after OVERALL VERDICT
        ov_idx = body_text.find("OVERALL VERDICT")
        snippet = body_text[ov_idx:ov_idx + 100]
        if "FAILED" in snippet:
            err = _get_error_text(driver)
            logger.info(f"Verdict: FAILED - {err}")
            return "FAILED", err
        elif "SUCCESS" in snippet:
            logger.info("Verdict: SUCCESS")
            return "SUCCESS", ""

    if any(kw in body_text for kw in ["FAILED", "FAILURE", "ERROR", "INVALID", "REJECTED"]):
        err = _get_error_text(driver)
        logger.info(f"Verdict: FAILED - {err}")
        return "FAILED", err

    if "SUCCESS" in body_text:
        logger.info("Verdict: SUCCESS")
        return "SUCCESS", ""

    logger.warning("Verdict: UNKNOWN (no clear SUCCESS/FAILED found)")
    return "UNKNOWN", "Could not determine verdict"


def _get_pipeline_error(driver: webdriver.Chrome) -> str:
    """
    Extract the full pipeline failure reason text if present.

    Looks for known detailed-reason phrases first, and for each match walks
    up several ancestor levels, keeping the longest (most detailed)
    surrounding text rather than the first/shallowest hit.
    """
    # Checked in order of specificity: exact reason phrases first, generic
    # heading phrases last (only used as a fallback if nothing else matches).
    trigger_phrases = [
        "PDF validation failed",
        "DOCUMENT_TYPE_MISMATCH",
        "Document type mismatch",
        "PDF_MULTIPLE_PAGES",
        "Processing could not be completed",
        "PIPELINE FAILURE",
        "Pipeline failure",
    ]

    for phrase in trigger_phrases:
        try:
            els = driver.find_elements(By.XPATH, f"//*[contains(text(),'{phrase}')]")
        except Exception:
            continue

        best_text = ""
        for el in els:
            if not el.is_displayed():
                continue

            # Walk up multiple ancestor levels and keep the longest text
            # found — the more detailed reason box will have more content
            # than just the short heading line.
            candidates = [el]
            try:
                candidates += el.find_elements(By.XPATH, "ancestor::*[position()<=6]")
            except Exception:
                pass

            for cand in candidates:
                try:
                    text = cand.text.strip().replace("\n", " — ")
                except Exception:
                    continue
                if text and 0 < len(text) < 800 and len(text) > len(best_text):
                    best_text = text

        if best_text:
            logger.info(f"Pipeline failure reason captured: {best_text}")
            return best_text[:500]

    return ""


def _get_error_text(driver: webdriver.Chrome) -> str:
    """Try to extract a specific error message from the page."""
    # First try pipeline error
    pe = _get_pipeline_error(driver)
    if pe:
        return pe

    candidates = [
        "//*[contains(@class,'error') or contains(@class,'alert') or contains(@class,'danger')]",
        "//*[contains(text(),'Error') or contains(text(),'Failed') or contains(text(),'Invalid') or contains(text(),'validation')]",
    ]
    for xp in candidates:
        try:
            els = driver.find_elements(By.XPATH, xp)
            for el in els:
                if el.is_displayed():
                    t = el.text.strip().replace("\n", " ")
                    if t and 5 < len(t) < 300:
                        return t
        except Exception:
            pass

    return "Verification failed on portal"


# ─────────────────────────────────────────────────────────────
# Scrape all label-value pairs from an already-open Extraction tab
# ─────────────────────────────────────────────────────────────
def _scrape_fields_only(driver: webdriver.Chrome) -> Dict[str, str]:
    """
    Runs the parsing strategies only — assumes the Extraction tab is
    already open (the caller is responsible for clicking it). Kept
    separate from get_extracted_fields() so the fast success path in
    parse_result() doesn't have to click the tab twice.
    """
    for parser in (_parse_via_javascript, _parse_table, _parse_div_rows, _parse_dl, _parse_generic_text):
        fields = parser(driver)
        if fields:
            logger.info(f"{parser.__name__}: {len(fields)} fields")
            return fields

    logger.warning("No fields could be extracted.")
    return {}


def get_extracted_fields(driver: webdriver.Chrome) -> Dict[str, str]:
    """
    Click the Extraction tab and dynamically scrape all label-value rows.
    Kept for any caller that needs the "click tab + scrape" behavior in one
    step. parse_result() below uses _scrape_fields_only() directly instead,
    since it has usually already clicked the tab as part of its fast path.
    """
    logger.info("Clicking Extraction tab...")
    try:
        ext_tab = _find_element(driver, XPATHS["extraction_tab"], timeout=10)
        _click_safe(driver, ext_tab)
        time.sleep(0.5)
        logger.info("On Extraction tab.")
    except NoSuchElementException:
        logger.warning("Extraction tab not found — attempting to read current page.")

    return _scrape_fields_only(driver)


# ─────────────────────────────────────────────────────────────
# Strategy 1 — JavaScript extraction (most reliable for SPAs)
# ─────────────────────────────────────────────────────────────
def _parse_via_javascript(driver: webdriver.Chrome) -> Dict[str, str]:
    """
    Use JS to find all element pairs that look like label | value rows.
    Handles any DOM structure robustly.
    """
    script = """
    var result = {};

    // --- Try rows inside tables ---
    var rows = document.querySelectorAll('table tr');
    rows.forEach(function(row) {
        var cells = row.querySelectorAll('td');
        if (cells.length >= 2) {
            var label = cells[0].innerText.trim().replace(/:$/, '');
            var value = cells[1].innerText.trim();
            if (label && label.length < 80) {
                result[label] = value;
            }
        }
    });
    if (Object.keys(result).length > 0) return result;

    // --- Try any element with two child spans/divs (label + value pattern) ---
    var containers = document.querySelectorAll(
        '[class*="field"], [class*="row"], [class*="item"], [class*="extract"]'
    );
    containers.forEach(function(c) {
        var children = Array.from(c.children).filter(function(ch) {
            return ch.innerText && ch.innerText.trim();
        });
        if (children.length === 2) {
            var label = children[0].innerText.trim().replace(/:$/, '');
            var value = children[1].innerText.trim();
            if (label && label.length < 80 && !result[label]) {
                result[label] = value;
            }
        }
    });
    if (Object.keys(result).length > 0) return result;

    // --- Try definition lists ---
    var dts = document.querySelectorAll('dt');
    var dds = document.querySelectorAll('dd');
    for (var i = 0; i < dts.length; i++) {
        var label = dts[i].innerText.trim().replace(/:$/, '');
        var value = dds[i] ? dds[i].innerText.trim() : '';
        if (label) result[label] = value;
    }
    return result;
    """
    try:
        data = driver.execute_script(script)
        if isinstance(data, dict) and data:
            return {str(k): str(v) for k, v in data.items() if k and str(k).strip()}
    except Exception as e:
        logger.debug(f"JS parse error: {e}")
    return {}


# ─────────────────────────────────────────────────────────────
# Strategy 2 — HTML table rows
# ─────────────────────────────────────────────────────────────
def _parse_table(driver: webdriver.Chrome) -> Dict[str, str]:
    fields = {}
    try:
        rows = driver.find_elements(By.XPATH, "//table//tr")
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 2:
                label = cells[0].text.strip().rstrip(":")
                value = cells[1].text.strip()
                if label and len(label) < 80:
                    fields[label] = value
    except Exception as e:
        logger.debug(f"Table parse: {e}")
    return fields


# ─────────────────────────────────────────────────────────────
# Strategy 3 — Div/span row pairs with label+value classes
# ─────────────────────────────────────────────────────────────
def _parse_div_rows(driver: webdriver.Chrome) -> Dict[str, str]:
    fields = {}
    label_xpaths = [
        "//*[contains(@class,'label') or contains(@class,'key') or contains(@class,'field-name') or contains(@class,'name')]",
    ]
    for lxp in label_xpaths:
        try:
            label_els = driver.find_elements(By.XPATH, lxp)
            for lel in label_els:
                label_text = lel.text.strip().rstrip(":")
                if not label_text or len(label_text) > 80:
                    continue
                try:
                    # Look for sibling or next element as value
                    val_el = lel.find_element(
                        By.XPATH,
                        "./following-sibling::*[1]"
                    )
                    value = val_el.text.strip()
                    if value:
                        fields[label_text] = value
                except NoSuchElementException:
                    pass
            if fields:
                return fields
        except Exception as e:
            logger.debug(f"Div-row parse: {e}")
    return fields


# ─────────────────────────────────────────────────────────────
# Strategy 4 — Definition lists <dl><dt><dd>
# ─────────────────────────────────────────────────────────────
def _parse_dl(driver: webdriver.Chrome) -> Dict[str, str]:
    fields = {}
    try:
        dts = driver.find_elements(By.TAG_NAME, "dt")
        dds = driver.find_elements(By.TAG_NAME, "dd")
        for dt, dd in zip(dts, dds):
            label = dt.text.strip().rstrip(":")
            value = dd.text.strip()
            if label:
                fields[label] = value
    except Exception as e:
        logger.debug(f"DL parse: {e}")
    return fields


# ─────────────────────────────────────────────────────────────
# Strategy 5 — Generic alternating line text scan
# ─────────────────────────────────────────────────────────────
def _parse_generic_text(driver: webdriver.Chrome) -> Dict[str, str]:
    """
    Last resort: look for a container element with text that alternates
    label / value on adjacent lines.
    """
    fields = {}
    try:
        containers = driver.find_elements(
            By.XPATH,
            "//*[contains(@class,'extract') or contains(@class,'result') "
            "or contains(@class,'fields')]"
        )
        for container in containers:
            lines = [l.strip() for l in container.text.split("\n") if l.strip()]
            if len(lines) >= 4:
                for i in range(0, len(lines) - 1, 2):
                    label = lines[i].rstrip(":")
                    value = lines[i + 1] if i + 1 < len(lines) else ""
                    # Only accept short labels (field names are typically < 40 chars)
                    if label and len(label) < 50:
                        fields[label] = value
            if fields:
                return fields
    except Exception as e:
        logger.debug(f"Generic text parse: {e}")
    return fields


# ─────────────────────────────────────────────────────────────
# Full result collection (called from main.py)
# ─────────────────────────────────────────────────────────────
def parse_result(
    driver: webdriver.Chrome,
    filename: str,
    doc_type: str,
    processing_time: float,
    already_on_extraction_tab: bool = False,
) -> Dict[str, str]:
    """
    Collect the full result for one document.

    Args:
        already_on_extraction_tab:
            True  — wait_for_result() already found and clicked the
                    Extraction tab. Skip the element search and scrape
                    immediately (SUCCESS fast path).
            False — result page appeared but no Extraction tab was present
                    (FAILED / UNKNOWN). Read the Overview tab for the
                    failure reason.

    This keeps parse_result() in sync with wait_for_result()'s new
    extraction-tab-first detection: for SUCCESS documents there is no
    longer a duplicate "find Extraction tab" scan after the wait.
    """
    if already_on_extraction_tab:
        # Fast path — Extraction tab is already open, scrape directly.
        logger.info("Scraping Extraction tab (already open from wait_for_result)...")
        fields = _scrape_fields_only(driver)
        take_screenshot(driver, f"extracted_{filename}")
        return {
            "Filename":          filename,
            "Document Type":     doc_type,
            "Status":            "SUCCESS",
            "Error Message":     "",
            "_processing_time":  round(processing_time, 1),
            **fields,
        }

    # Slow path — no Extraction tab on result page → read failure reason.
    logger.info("Extraction tab not present — reading Overview for failure reason.")
    take_screenshot(driver, f"result_{filename}")
    status, error_msg = get_status(driver)
    return {
        "Filename":         filename,
        "Document Type":    doc_type,
        "Status":           status,
        "Error Message":    error_msg,
        "_processing_time": round(processing_time, 1),
    }
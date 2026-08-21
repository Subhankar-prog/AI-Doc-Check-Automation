"""
doc_uploader.py — Core upload, submit, and wait logic for VERIFAI portal
"""
import time
import logging
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, ElementClickInterceptedException,
    StaleElementReferenceException, WebDriverException
)

from config import (
    PORTAL_URL, XPATHS, ENGINE_TIMEOUT, POLL_INTERVAL,
    RETRY_ON_TIMEOUT, SCREENSHOTS_DIR, DOCUMENT_TYPES
)

logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────────────────
# Helper: find element by multiple XPath fallbacks
# ──────────────────────────────────────────────────────────
def _find_element(driver: webdriver.Chrome, *xpaths: str, timeout: int = 10):
    """Try each XPath in order; return first found element within timeout."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        for xpath in xpaths:
            try:
                el = driver.find_element(By.XPATH, xpath)
                if el.is_displayed():
                    return el
            except NoSuchElementException:
                pass
        time.sleep(0.5)
    raise NoSuchElementException(
        f"None of the XPaths found within {timeout}s: {xpaths}"
    )


def _click_safe(driver: webdriver.Chrome, element):
    """Click element, scroll into view first; fallback to JS click."""
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    time.sleep(0.3)
    try:
        element.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", element)


def take_screenshot(driver: webdriver.Chrome, name: str) -> str:
    """Save a screenshot and return its absolute path."""
    SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)
    ts = time.strftime("%Y%m%d_%H%M%S")
    path = SCREENSHOTS_DIR / f"{ts}_{name}.png"
    driver.save_screenshot(str(path))
    logger.info(f"Screenshot saved: {path}")
    return str(path)


# ──────────────────────────────────────────────────────────
# Step 1: Navigate to the Verify Document upload page
# ──────────────────────────────────────────────────────────
def navigate_to_upload(driver: webdriver.Chrome):
    """Click 'Verify Document' in the sidebar to reach the upload form."""
    logger.info("Navigating to 'Verify Document' page...")
    try:
        btn = _find_element(driver, XPATHS["verify_document_link"], timeout=10)
        _click_safe(driver, btn)
        time.sleep(1.5)
        logger.info("On Verify Document page.")
    except NoSuchElementException:
        # If already on the right page (after New Submission), skip
        logger.info("'Verify Document' link not found — may already be on form page.")


# ──────────────────────────────────────────────────────────
# Step 2: Select document type from dropdown
# ──────────────────────────────────────────────────────────
def select_document_type(driver: webdriver.Chrome, doc_type: str):
    """
    Select the given document type from the dropdown.
    doc_type must match a key or value in DOCUMENT_TYPES config.
    """
    # Resolve to portal dropdown display text
    display_text = DOCUMENT_TYPES.get(doc_type, doc_type)
    logger.info(f"Selecting document type: '{display_text}'")

    # Find the <select> element
    try:
        select_el = _find_element(
            driver,
            XPATHS["doc_type_select"],
            XPATHS["doc_type_select"],
            timeout=10
        )
        sel = Select(select_el)

        # Try selecting by visible text first
        try:
            sel.select_by_visible_text(display_text)
            logger.info(f"Selected by visible text: '{display_text}'")
            return
        except Exception:
            pass

        # Try partial text match across all options
        for option in sel.options:
            if display_text.lower() in option.text.lower():
                sel.select_by_visible_text(option.text)
                logger.info(f"Selected by partial match: '{option.text}'")
                return

        raise ValueError(
            f"Document type '{display_text}' not found in dropdown. "
            f"Available: {[o.text for o in sel.options]}"
        )

    except NoSuchElementException:
        # Might be a custom (non-native) dropdown — handle via click
        logger.warning("Native <select> not found. Trying custom dropdown...")
        _select_custom_dropdown(driver, display_text)


def _select_custom_dropdown(driver: webdriver.Chrome, display_text: str):
    """Handle custom (non-native) styled dropdowns."""
    # Click the dropdown trigger
    trigger = driver.find_element(
        By.XPATH,
        "//div[contains(@class,'select') or contains(@class,'dropdown')][@role='combobox' or contains(@class,'trigger')]"
    )
    _click_safe(driver, trigger)
    time.sleep(0.5)

    # Find the option with matching text
    option = driver.find_element(
        By.XPATH,
        f"//*[@role='option' or contains(@class,'option')][contains(text(),'{display_text}')]"
    )
    _click_safe(driver, option)
    logger.info(f"Custom dropdown: selected '{display_text}'")


# ──────────────────────────────────────────────────────────
# Step 3: Verify Plain Extraction radio is selected
# ──────────────────────────────────────────────────────────
def ensure_plain_extraction(driver: webdriver.Chrome):
    """Make sure 'Plain extraction' radio button is selected."""
    logger.info("Ensuring 'Plain extraction' is selected...")

    # Strategy 1: click the label text element
    try:
        el = _find_element(
            driver,
            XPATHS["plain_extraction_btn"],
            XPATHS["plain_extraction_btn_v2"],
            timeout=8
        )
        # Check if parent radio input is already checked
        try:
            radio_input = el.find_element(
                By.XPATH,
                ".//ancestor::label//input[@type='radio'] | .//preceding-sibling::input[@type='radio'] | .//following-sibling::input[@type='radio']"
            )
            if not radio_input.is_selected():
                _click_safe(driver, el)
                logger.info("Clicked Plain extraction label.")
            else:
                logger.info("Plain extraction already selected.")
        except NoSuchElementException:
            _click_safe(driver, el)
            logger.info("Clicked Plain extraction text element.")
        return
    except NoSuchElementException:
        pass

    # Strategy 2: find the radio input directly
    try:
        radio = _find_element(driver, XPATHS["plain_extraction_btn_v2"], timeout=5)
        if not radio.is_selected():
            _click_safe(driver, radio)
        logger.info("Plain extraction radio input clicked.")
    except NoSuchElementException:
        logger.warning("Plain extraction radio not found — may already be default.")


# ──────────────────────────────────────────────────────────
# Step 4: Upload the document file
# ──────────────────────────────────────────────────────────
def upload_file(driver: webdriver.Chrome, file_path: str):
    """
    Upload a file to the portal.

    The portal's <input type='file'> is hidden — Selenium's normal
    find_element() skips hidden elements. We use JavaScript to locate it
    directly, make it interactable, then send the file path.

    Also handles the case where an 'Upload' button must be clicked first
    to trigger the file input to appear.
    """
    abs_path = str(Path(file_path).resolve())
    logger.info(f"Uploading file: {abs_path}")

    if not Path(abs_path).exists():
        raise FileNotFoundError(f"File not found: {abs_path}")

    # Strategy 1: Use JS to find ALL file inputs (including hidden ones)
    file_input = _get_file_input_via_js(driver)

    if file_input is None:
        # Strategy 2: Click "Upload" or "Choose File" button first, then find input
        logger.info("File input not found — trying to click Upload button first...")
        _click_upload_trigger(driver)
        time.sleep(1.5)
        file_input = _get_file_input_via_js(driver)

    if file_input is None:
        raise NoSuchElementException(
            "Could not locate <input type='file'> even after clicking upload trigger.\n"
            "Check the portal's file upload mechanism in DevTools."
        )

    # Make the input fully interactable via JS (remove hidden/display:none)
    driver.execute_script(
        """
        arguments[0].style.display    = 'block';
        arguments[0].style.visibility = 'visible';
        arguments[0].style.opacity    = '1';
        arguments[0].style.position   = 'fixed';
        arguments[0].style.top        = '0';
        arguments[0].style.left       = '0';
        arguments[0].style.zIndex     = '9999';
        arguments[0].removeAttribute('hidden');
        """,
        file_input
    )
    time.sleep(0.3)
    file_input.send_keys(abs_path)
    time.sleep(2)  # wait for preview/thumbnail to render

    # Check if portal rejected the file (size / format / multi-page error)
    # Screenshot 2 shows a pink banner: "Upload a JPEG, PNG, WebP, or PDF document: up to 9 MB."
    error_msg = _check_upload_error(driver)
    if error_msg:
        raise UploadRejectedError(error_msg)

    logger.info("File uploaded successfully.")


class UploadRejectedError(Exception):
    """Raised when the portal rejects a file at upload time (wrong format, too large, etc.)"""
    pass


def _check_upload_error(driver: webdriver.Chrome) -> str:
    """
    Check if the portal shows an upload error banner (size/format rejection).
    Screenshot shows pink alert banner below Submit button:
    'Upload a JPEG, PNG, WebP, or PDF document: up to 5 MB.'
    """
    # 1. Use JavaScript to inspect visible text for size/format error banners
    js_check = """
    var banners = document.querySelectorAll('*');
    for (var i = 0; i < banners.length; i++) {
        var el = banners[i];
        if (el.children.length === 0 && el.innerText) {
            var txt = el.innerText.trim();
            // Check for upload error text
            if ((txt.includes('up to 5 MB') || txt.includes('up to 5MB') ||
                 txt.includes('document: up to') || txt.includes('MB.') ||
                 txt.includes('Only single-page') || txt.includes('validation failed')) &&
                el.offsetWidth > 0 && el.offsetHeight > 0) {
                return txt;
            }
        }
    }
    // Also check elements with alert/error/danger classes
    var errs = document.querySelectorAll('[class*="error"], [class*="alert"], [class*="danger"], [class*="invalid"]');
    for (var j = 0; j < errs.length; j++) {
        var e = errs[j];
        if (e.offsetWidth > 0 && e.offsetHeight > 0 && e.innerText && e.innerText.trim().length > 5) {
            return e.innerText.trim();
        }
    }
    return '';
    """
    try:
        err = driver.execute_script(js_check)
        if err and isinstance(err, str) and len(err.strip()) > 3:
            logger.warning(f"Upload error detected via JS: {err.strip()}")
            return err.strip()
    except Exception as e:
        logger.debug(f"JS upload error check: {e}")

    # 2. Selenium XPath fallback
    error_xpaths = [
        "//*[contains(text(),'up to 5 MB') or contains(text(),'up to 5MB') or contains(text(),'document: up to')]",
        "//*[contains(@class,'error') or contains(@class,'alert') or contains(@class,'danger')][not(ancestor::*[contains(@class,'result')])]",
    ]
    for xp in error_xpaths:
        try:
            els = driver.find_elements(By.XPATH, xp)
            for el in els:
                if el.is_displayed():
                    text = el.text.strip()
                    if text and len(text) < 300:
                        logger.warning(f"Upload error detected via XPath: {text}")
                        return text
        except Exception:
            pass
    return ""




def _get_file_input_via_js(driver: webdriver.Chrome):
    """Use JavaScript to find the first <input type='file'> including hidden ones."""
    try:
        el = driver.execute_script(
            "return document.querySelector('input[type=\"file\"]');"
        )
        if el:
            logger.info("File input found via JS querySelector.")
            return el
    except Exception as e:
        logger.debug(f"JS querySelector failed: {e}")

    # Try finding all inputs and filter by type
    try:
        els = driver.execute_script(
            "return Array.from(document.querySelectorAll('input')).filter(i => i.type === 'file');"
        )
        if els:
            logger.info(f"File input found via JS querySelectorAll (count={len(els)}).")
            return els[0]
    except Exception as e:
        logger.debug(f"JS querySelectorAll failed: {e}")

    return None


def _click_upload_trigger(driver: webdriver.Chrome):
    """Click a button/label that triggers the file input to appear."""
    upload_triggers = [
        "//*[contains(text(),'Upload') or contains(text(),'upload')]"
        "[not(contains(text(),'Submit'))]",
        "//label[contains(@for,'file') or contains(@class,'upload')]",
        "//button[contains(@class,'upload') or contains(@class,'file')]",
        "//*[contains(@class,'upload-btn') or contains(@class,'file-btn')]",
    ]
    for xp in upload_triggers:
        try:
            btn = driver.find_element(By.XPATH, xp)
            if btn.is_displayed():
                _click_safe(driver, btn)
                logger.info(f"Clicked upload trigger: {btn.text.strip()[:50]}")
                return
        except NoSuchElementException:
            pass
    logger.warning("No upload trigger button found — file input may already be present.")



# ──────────────────────────────────────────────────────────
# Step 5: Verify state, then click Submit to Engine
# ──────────────────────────────────────────────────────────
def verify_and_submit(driver: webdriver.Chrome, doc_type: str):
    """
    Final pre-submit checks:
      - Dropdown shows correct doc type
      - Plain extraction is selected
      - Check for upload rejection error banners
    Then click Submit to engine.
    """
    logger.info("Pre-submit verification...")

    # Check for upload rejection error banner (e.g. file size > 5MB, format error)
    upload_err = _check_upload_error(driver)
    if upload_err:
        raise UploadRejectedError(upload_err)

    # Check dropdown value
    try:
        select_el = _find_element(
            driver,
            XPATHS["doc_type_select"],
            XPATHS["doc_type_select"],
            timeout=5
        )
        sel = Select(select_el)
        current = sel.first_selected_option.text.strip()
        expected = DOCUMENT_TYPES.get(doc_type, doc_type)
        if expected.lower() not in current.lower():
            logger.warning(
                f"Dropdown mismatch! Expected '{expected}', found '{current}'. Re-selecting."
            )
            select_document_type(driver, doc_type)
    except Exception as e:
        logger.warning(f"Could not verify dropdown: {e}")

    # Check plain extraction
    ensure_plain_extraction(driver)

    # Re-check upload error before submit
    upload_err = _check_upload_error(driver)
    if upload_err:
        raise UploadRejectedError(upload_err)

    # Click Submit
    logger.info("Clicking 'Submit to engine'...")
    submit_btn = _find_element(
        driver,
        XPATHS["submit_to_engine"],
        XPATHS["submit_to_engine_v2"],
        timeout=10
    )
    _click_safe(driver, submit_btn)
    time.sleep(1.0)

    # Post-click check: did an immediate validation error popup?
    post_err = _check_upload_error(driver)
    if post_err:
        raise UploadRejectedError(post_err)

    logger.info("Submitted to engine. Waiting for result...")


# ──────────────────────────────────────────────────────────
# Step 6: Wait for engine to finish
# ──────────────────────────────────────────────────────────
def wait_for_result(driver: webdriver.Chrome) -> bool:
    """
    Wait until the engine finishes processing.
    Detected by: "TRANSACTION COMPLETE" heading, pipeline failure, or result-tabs appearing.
    Returns True if result page appeared, False on timeout.
    """
    logger.info(f"Waiting for engine result (timeout={ENGINE_TIMEOUT}s)...")
    start    = time.time()
    deadline = start + ENGINE_TIMEOUT

    while time.time() < deadline:
        elapsed = int(time.time() - start)

        # 1. Check if the result page has appeared FIRST
        # "TRANSACTION COMPLETE" or "Pipeline failure result" or "Plain extraction result" or result-tabs div or New submission button
        try:
            el = driver.find_element(By.XPATH, XPATHS["result_appeared"])
            if el.is_displayed():
                logger.info(f"Result appeared after {elapsed}s.")
                time.sleep(1)   # small stabilisation pause
                return True
        except (NoSuchElementException, StaleElementReferenceException):
            pass

        # 2. Fallback: check body text for fast detection
        try:
            body = driver.find_element(By.TAG_NAME, "body").text.upper()
            if any(k in body for k in ["TRANSACTION COMPLETE", "PIPELINE FAILURE", "OVERALL VERDICT", "NEW SUBMISSION"]):
                logger.info(f"Result keyword detected in body after {elapsed}s.")
                time.sleep(1)
                return True
        except Exception:
            pass

        # 3. Check if upload error banner is on screen (e.g. file size > 5 MB rejection)
        upload_err = _check_upload_error(driver)
        if upload_err:
            logger.warning(f"Upload rejection banner detected during wait: {upload_err}")
            raise UploadRejectedError(upload_err)

        # 4. Check if still in processing state
        try:
            proc = driver.find_element(By.XPATH, XPATHS["processing_indicator"])
            if proc.is_displayed():
                logger.info(f"  [{elapsed}s] Engine processing...")
                time.sleep(POLL_INTERVAL)
                continue
        except (NoSuchElementException, StaleElementReferenceException):
            pass

        time.sleep(POLL_INTERVAL)

    logger.warning(f"Engine timeout after {ENGINE_TIMEOUT}s.")
    return False




# ──────────────────────────────────────────────────────────
# Step 7: Click New Submission button
# ──────────────────────────────────────────────────────────
def click_new_submission(driver: webdriver.Chrome):
    """
    Click 'New submission' button to reset the form for the next document.
    If already on the submission/upload form (or if button not found),
    safely navigates/refreshes to the Verify Document page.
    """
    logger.info("Resetting form for next submission...")
    try:
        btn = _find_element(
            driver,
            XPATHS["new_submission"],
            XPATHS["new_submission_v2"],
            timeout=3
        )
        _click_safe(driver, btn)
        time.sleep(1.5)
        logger.info("New submission clicked — form reset.")
        return
    except NoSuchElementException:
        pass

    # If 'New submission' button wasn't found, navigate directly to upload page
    logger.info("Navigating to Verify Document page to reset form.")
    navigate_to_upload(driver)


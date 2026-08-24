"""
driver_setup.py - Chrome WebDriver setup using Selenium's built-in manager

Uses selenium-manager (built into Selenium 4.6+) which automatically downloads
the correct chromedriver version for your OS architecture (win64/win32/mac/linux).
No external webdriver-manager package needed.
"""
import logging
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from config import HEADLESS_MODE, PAGE_LOAD_TIMEOUT

logger = logging.getLogger(__name__)


def create_driver() -> webdriver.Chrome:
    """
    Create and return a configured Chrome WebDriver.
    Selenium Manager (built-in) automatically fetches the correct
    win64 chromedriver — no manual driver management needed.
    """
    options = Options()

    if HEADLESS_MODE:
        options.add_argument("--headless=new")
        # --start-maximized doesn't reliably apply in headless mode, so set
        # an explicit desktop-sized window to avoid the portal rendering
        # differently (e.g. a smaller/mobile layout) than in normal mode.
        options.add_argument("--window-size=1920,1080")

    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-infobars")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--log-level=3")  # Suppress internal Chrome logs (GCM, DevTools, WebGL)
    options.add_argument("--silent")

    # Prevent Chrome from asking where to save downloaded files
    options.add_experimental_option("prefs", {
        "download.default_directory": "",
        "download.prompt_for_download": False,
    })

    # Suppress "Chrome is being controlled by automated software" bar & internal Chrome log spam
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    options.add_experimental_option("useAutomationExtension", False)


    # Let Selenium Manager handle the driver automatically (correct architecture)
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    driver.implicitly_wait(5)

    logger.info("Chrome WebDriver started successfully.")
    return driver
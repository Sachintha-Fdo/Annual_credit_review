import os
import time
import logging
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from dotenv import load_dotenv

# Load credentials from .env file
load_dotenv("credentials.env")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(levelname)s] %(asctime)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# --- CREDENTIALS SETUP ---
# Store credentials in credentials.env for security.
ERP_USERNAME = os.environ.get('ERP_USERNAME')
ERP_PASSWORD = os.environ.get('ERP_PASSWORD')

if not ERP_USERNAME or not ERP_PASSWORD:
    logging.error("ERP_USERNAME and ERP_PASSWORD must be set in credentials.env.")
    exit(1)

ERP_URL = os.environ.get('ERP_URL')

# Set download directory to 'data' folder
DATA_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), 'data'))
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)


def setup_firefox_driver():
    options = Options()
    options.set_preference("webdriver_accept_untrusted_certs", True)
    options.set_preference("acceptInsecureCerts", True)
    # Set download preferences
    options.set_preference("browser.download.folderList", 2)  # custom location
    options.set_preference("browser.download.dir", DATA_DIR)
    options.set_preference(
        "browser.helperApps.neverAsk.saveToDisk",
        "application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream"
    )
    options.set_preference("pdfjs.disabled", True)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.download.useDownloadDir", True)
    options.set_preference("browser.helperApps.alwaysAsk.force", False)
    # options.headless = True  # Uncomment for headless mode
    driver = webdriver.Firefox(options=options)
    driver.maximize_window()
    return driver


def login(driver, username, password):
    logging.info("Navigating to login page")
    driver.get(ERP_URL)
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "txtusername"))
        )
        driver.find_element(By.ID, "txtusername").send_keys(username)
        driver.find_element(By.ID, "txtpassword").send_keys(password)
        driver.find_element(By.CLASS_NAME, "login100-form-btn").click()
        logging.info("Login submitted")
    except Exception as e:
        logging.error(f"Login failed: {e}")
        raise


def expand_menu(driver, menu_text):
    logging.info(f"Expanding menu: {menu_text}")
    dropdown_buttons = driver.find_elements(By.CLASS_NAME, "dropdown-btn")
    for btn in dropdown_buttons:
        try:
            span = btn.find_element(By.TAG_NAME, "span")
            if span.text.strip() == menu_text:
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(1)
                return True
        except Exception:
            continue
    logging.warning(f"Menu '{menu_text}' not found.")
    return False


def click_reports(driver):
    logging.info("Clicking 'Reports' link")
    wait = WebDriverWait(driver, 10)
    try:
        reporting_link = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(text(), 'Reports')]")
        ))
        driver.execute_script("arguments[0].click();", reporting_link)
        logging.info("Navigated to Reports.")
    except TimeoutException:
        logging.error("'Reports' link not found or not clickable.")
        raise


def switch_to_report_iframe(driver):
    time.sleep(2)
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    logging.info(f"Found {len(iframes)} iframe(s)")
    for index, iframe in enumerate(iframes):
        try:
            driver.switch_to.frame(iframe)
            driver.find_element(By.ID, "report_type")
            logging.info(f"Found 'report_type' dropdown inside iframe index {index}")
            return True
        except NoSuchElementException:
            driver.switch_to.default_content()
    # Try on main page as fallback
    try:
        driver.find_element(By.ID, "report_type")
        logging.info("Found 'report_type' dropdown on main page")
        return True
    except NoSuchElementException:
        logging.error("Could not find 'report_type' dropdown in any iframe or main page.")
        return False


def select_report_type(driver, report_name):
    try:
        select = Select(driver.find_element(By.ID, "report_type"))
        select.select_by_visible_text(report_name)
        logging.info(f"Selected '{report_name}' from dropdown.")
    except Exception as e:
        logging.error(f"Dropdown selection failed: {e}")
        raise


def set_from_date(driver, date_str):
    try:
        from_date_input = driver.find_element(By.ID, "from_date")
        driver.execute_script("arguments[0].value = arguments[1];", from_date_input, date_str)
        logging.info(f"Set 'from_date' to {date_str}")
    except Exception as e:
        logging.error(f"Failed to set 'from_date': {e}")
        raise


def export_to_excel(driver):
    try:
        export_btn = driver.find_element(By.ID, "btn_export")
        driver.execute_script("arguments[0].click();", export_btn)
        logging.info("Clicked 'Excel' export button.")
    except Exception as e:
        logging.error(f"Failed to click Excel export button: {e}")
        raise


def wait_for_download(download_dir, timeout=100):
    """
    Waits for a new .xlsx file to appear in the download directory.
    Returns the filename if found, raises Exception if timeout.
    """
    seconds = 0
    before = set(os.listdir(download_dir))
    while seconds < timeout:
        after = set(os.listdir(download_dir))
        new_files = after - before
        for f in new_files:
            if f.endswith('.xlsx'):
                file_path = os.path.join(download_dir, f)
                if not file_path.endswith('.crdownload'):
                    return f
        time.sleep(1)
        seconds += 1
    raise Exception("Download did not complete in time.")


def main():
    driver = setup_firefox_driver()
    try:
        login(driver, ERP_USERNAME, ERP_PASSWORD)
        # Wait for menu to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "dropdown-btn"))
        )
        time.sleep(2)
        expand_menu(driver, "Operation")
        expand_menu(driver, "Annual Credit Review")
        click_reports(driver)
        if not switch_to_report_iframe(driver):
            raise Exception("Could not find 'report_type' dropdown.")
        select_report_type(driver, "Annual Credit Review Evaluation report")
        yesterday = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
        set_from_date(driver, yesterday)
        export_to_excel(driver)
        downloaded_file = wait_for_download(DATA_DIR, timeout=100)
        logging.info(f"[SUCCESS] Report export process completed. File saved as: {downloaded_file} in {DATA_DIR}")
    except Exception as e:
        logging.error(f"[ERROR] {str(e)}")
    finally:
        driver.quit()


if __name__ == "__main__":
    main() 
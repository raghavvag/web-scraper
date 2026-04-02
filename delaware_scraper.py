import time
from typing import Dict, List

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

BASE_URL = "https://icis.corp.delaware.gov/ecorp/entitysearch/namesearch.aspx"
DEFAULT_DELAY_SECONDS = 8


def setup_driver(headless: bool = False) -> webdriver.Chrome:
    """Create and return a Chrome WebDriver instance."""
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(60)
    return driver


def search_entity(driver: webdriver.Chrome, file_number: str, wait_seconds: int = 25) -> None:
    """Open search page, fill file number, and submit using Search button (fallback Enter)."""
    driver.get(BASE_URL)

    wait = WebDriverWait(driver, wait_seconds)
    file_input = wait.until(
        EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_frmFileNumber"))
    )

    file_input.clear()
    file_input.send_keys(str(file_number))
    # Small pause before submit to mimic normal user pacing.
    time.sleep(1)

    # Prefer clicking Search like a real user; fallback to Enter.
    try:
        search_button = wait.until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_btnSubmit"))
        )
        if not search_button.is_enabled():
            driver.execute_script("arguments[0].removeAttribute('disabled');", search_button)
        search_button.click()
    except (TimeoutException, WebDriverException):
        file_input.send_keys(Keys.ENTER)


def extract_entity_name(driver: webdriver.Chrome, file_number: str, wait_seconds: int = 25) -> str:
    """Extract entity name from tblResults using file number match."""
    wait = WebDriverWait(driver, wait_seconds)

    try:
        wait.until(
            lambda d: d.find_elements(By.ID, "tblResults")
            or d.find_elements(By.XPATH, "//*[contains(., 'An error occurred while processing the request')]")
            or d.find_elements(By.XPATH, "//*[contains(., 'No Records Found') or contains(., 'No records found')]")
        )
    except TimeoutException:
        return "Not Found"

    if driver.find_elements(By.XPATH, "//*[contains(., 'An error occurred while processing the request')]"):
        return "Not Found"

    rows = driver.find_elements(By.XPATH, "//table[@id='tblResults']//tr")
    if len(rows) <= 1:
        return "Not Found"

    # Preferred: row where first column matches the requested file number.
    target_cells = driver.find_elements(
        By.XPATH,
        f"//table[@id='tblResults']//tr[td and contains(normalize-space(td[1]), '{file_number}')]/td[2]",
    )
    if target_cells:
        entity_name = target_cells[0].text.strip()
        if entity_name:
            return entity_name

    # Fallback: first data row in results table.
    first_data_cells = driver.find_elements(By.XPATH, "//table[@id='tblResults']//tr[td][1]/td")
    if len(first_data_cells) >= 2:
        entity_name = first_data_cells[1].text.strip()
        if entity_name:
            return entity_name

    return "Not Found"


def process_bulk(
    file_numbers: List[str],
    delay_seconds: float = DEFAULT_DELAY_SECONDS,
    max_retries: int = 2,
    headless: bool = False,
) -> List[Dict[str, str]]:
    """Search each file number and return file_number/entity_name records."""
    results: List[Dict[str, str]] = []
    driver = setup_driver(headless=headless)

    try:
        for i, file_number in enumerate(file_numbers):
            print(f"[INFO] Searching file number: {file_number}")
            entity_name = "Not Found"

            for attempt in range(1, max_retries + 1):
                try:
                    search_entity(driver, str(file_number))
                    entity_name = extract_entity_name(driver, str(file_number))

                    if entity_name != "Not Found":
                        print(f"[INFO] Found entity: {entity_name}")
                        break
                    else:
                        print(f"[WARN] Entity not found for {file_number} (attempt {attempt}/{max_retries})")
                        if attempt < max_retries:
                            time.sleep(2)
                except (TimeoutException, WebDriverException) as exc:
                    print(
                        f"[WARN] Attempt {attempt}/{max_retries} failed for {file_number}: {exc}"
                    )
                    if attempt == max_retries:
                        entity_name = "Not Found"
                    else:
                        time.sleep(2)

            results.append({"file_number": str(file_number), "entity_name": entity_name})

            if i < len(file_numbers) - 1:
                time.sleep(delay_seconds)
    finally:
        driver.quit()

    return results


def export_to_excel(results: List[Dict[str, str]], output_file: str = "output.xlsx") -> None:
    """Export file number and entity name results to Excel."""
    df = pd.DataFrame(results, columns=["file_number", "entity_name"])
    df.to_excel(output_file, index=False)


def main() -> None:
    file_numbers = ["10477355", "10477356", "10477357", "10477358", "10477359"]
    results = process_bulk(
        file_numbers=file_numbers,
        delay_seconds=DEFAULT_DELAY_SECONDS,
        max_retries=2,
        headless=False,
    )
    export_to_excel(results, output_file="output.xlsx")
    print(f"[INFO] Saved {len(results)} rows to output.xlsx")


if __name__ == "__main__":
    main()

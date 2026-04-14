from __future__ import annotations

import subprocess
import time
from pathlib import Path
import sys

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

ROOT = Path("/Users/shaananc/Library/CloudStorage/OneDrive-Personal/useful-scripts/arc_budget")
sys.path.insert(0, str(ROOT))
from RMSBudgetUploader import RMSBudgetBuilder

PROPOSAL_ID = "DP270101107"
ARC_PDF = Path(
    "/Users/shaananc/Library/CloudStorage/OneDrive-Personal/grants/arc-dp-2027/budget-justification/budget-arc.pdf"
)
NON_ARC_PDF = Path(
    "/Users/shaananc/Library/CloudStorage/OneDrive-Personal/grants/arc-dp-2027/budget-justification/budget-uom.pdf"
)

ARC_INPUT_ID = "346dff67-2f63-480c-b513-14865ed6706f__ead92b1d-3ccd-4a6f-821f-548191b7a9a5"
ARC_BUTTON_ID = "-delta-file-upload-button-346dff67-2f63-480c-b513-14865ed6706f__ead92b1d-3ccd-4a6f-821f-548191b7a9a5"
NON_ARC_INPUT_ID = "346dff67-2f63-480c-b513-14865ed6706f__a41fa76c-a89c-487d-8767-71586cd8a49b"
NON_ARC_BUTTON_ID = "-delta-file-upload-button-346dff67-2f63-480c-b513-14865ed6706f__a41fa76c-a89c-487d-8767-71586cd8a49b"


def resolve_matching_chromedriver() -> Path:
    chrome_binary = Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
    version_output = subprocess.check_output([str(chrome_binary), "--version"], text=True).strip()
    chrome_version = version_output.removeprefix("Google Chrome ").strip()
    cached_driver = (
        Path.home()
        / ".cache/selenium/chromedriver/mac-arm64"
        / chrome_version
        / "chromedriver"
    )
    if cached_driver.exists():
        return cached_driver
    raise FileNotFoundError(f"No matching cached ChromeDriver for Chrome {chrome_version}")


def login_interactive(builder: RMSBudgetBuilder) -> None:
    builder.login(ROOT / "rmscreds.txt")


def goto_budget_justification_page(builder: RMSBudgetBuilder) -> None:
    builder.driver.get(f"https://rms.arc.gov.au/RMS/Proposal/Form/Edit/{PROPOSAL_ID}")
    wait = WebDriverWait(builder.driver, 30)
    try:
        part_d_button = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button[data-part='D']"))
        )
        builder.driver.execute_script("arguments[0].click();", part_d_button)
        wait.until(
            lambda driver: "Project Cost"
            in driver.find_element(By.ID, "-delta-current-form-part").text
        )
        wait.until(EC.presence_of_element_located((By.ID, ARC_INPUT_ID)))
    except TimeoutException:
        debug_dir = ROOT / "tmp/web-debug"
        debug_dir.mkdir(parents=True, exist_ok=True)
        (debug_dir / "rms-budget-pdf-upload-timeout.html").write_text(builder.driver.page_source)
        builder.driver.save_screenshot(str(debug_dir / "rms-budget-pdf-upload-timeout.png"))
        raise
    builder.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", builder.driver.find_element(By.ID, ARC_INPUT_ID))
    time.sleep(1)


def upload_pdf(builder: RMSBudgetBuilder, input_id: str, button_id: str, pdf_path: Path, label: str) -> None:
    if not pdf_path.exists():
        raise FileNotFoundError(f"Missing PDF for {label}: {pdf_path}")

    wait = WebDriverWait(builder.driver, 30)
    input_el = wait.until(EC.presence_of_element_located((By.ID, input_id)))
    form_id = f"{input_id}-form"

    builder.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_el)
    input_el.send_keys(str(pdf_path))

    upload_button = builder.driver.find_element(By.ID, button_id)
    builder.driver.execute_script("arguments[0].click();", upload_button)

    def upload_complete(driver):
        text = driver.find_element(By.ID, form_id).text
        return ("No PDF uploaded" not in text) and (pdf_path.name in text or "pages" in text or "Uploaded at" in text)

    wait.until(upload_complete)
    print(f"Uploaded {label}: {pdf_path.name}")


def main() -> None:
    builder = RMSBudgetBuilder()
    options = ChromeOptions()
    options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    options.add_argument("--window-size=1440,1000")
    options.add_argument("--no-sandbox")
    builder.driver = webdriver.Chrome(
        service=ChromeService(executable_path=str(resolve_matching_chromedriver())),
        options=options,
    )

    login_interactive(builder)
    goto_budget_justification_page(builder)
    upload_pdf(builder, ARC_INPUT_ID, ARC_BUTTON_ID, ARC_PDF, "ARC budget justification")
    upload_pdf(
        builder,
        NON_ARC_INPUT_ID,
        NON_ARC_BUTTON_ID,
        NON_ARC_PDF,
        "non-ARC contributions",
    )

    debug_dir = ROOT / "tmp/web-debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    screenshot_path = debug_dir / "rms-budget-pdfs-uploaded.png"
    builder.driver.save_screenshot(str(screenshot_path))
    print(f"Saved screenshot to {screenshot_path}")


if __name__ == "__main__":
    main()

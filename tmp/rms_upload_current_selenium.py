from __future__ import annotations

import argparse
import json
import subprocess
import time
from pathlib import Path
import sys

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

ROOT = Path("/Users/shaananc/Library/CloudStorage/OneDrive-Personal/useful-scripts/arc_budget")
sys.path.insert(0, str(ROOT))
from RMSBudgetUploader import BudgetCategory, RMSBudgetBuilder

PAYLOAD = json.loads((ROOT / "tmp/rms_current_payload.json").read_text())


def canonical_name(name: str) -> str:
    return str(name).splitlines()[0].strip()


for year_data in PAYLOAD["years"]:
    for entry in year_data["entries"]:
        entry["name"] = canonical_name(entry["name"])


def visible_elements(driver, selector: str):
    elements = driver.find_elements(By.CSS_SELECTOR, selector)
    return [element for element in elements if element.is_displayed()]


def parse_args():
    parser = argparse.ArgumentParser(
        description="Sync the current RMS budget payload to RMS."
    )
    parser.add_argument(
        "--full-reset",
        action="store_true",
        help="Delete existing budget rows first, then re-add from scratch.",
    )
    parser.add_argument(
        "--no-prune",
        action="store_true",
        help="Do not delete stale managed rows that are absent from the payload.",
    )
    return parser.parse_args()


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


def login_interactive(builder: RMSBudgetBuilder):
    creds_lines = (ROOT / "rmscreds.txt").read_text().splitlines()
    credentials = {"username": creds_lines[0].strip(), "password": creds_lines[1].strip()}
    builder.driver.get("https://rms.arc.gov.au/RMS/ActionCentre")
    builder.driver.set_window_size(550, 691)

    if builder.driver.find_elements(By.LINK_TEXT, "Edit"):
        return

    if builder.driver.find_elements(By.ID, "emailAddress"):
        builder.driver.find_element(By.ID, "emailAddress").send_keys(
            credentials["username"]
        )
        password = builder.driver.find_element(By.ID, "password")
        password.click()
        password.send_keys(credentials["password"])
        try:
            builder.driver.find_element(By.ID, "login").click()
        except Exception:
            builder.driver.execute_script(
                "arguments[0].click();",
                builder.driver.find_element(By.ID, "login"),
            )

    deadline = time.time() + 10 * 60
    last_notice = 0.0
    while time.time() < deadline:
        if builder.driver.find_elements(By.LINK_TEXT, "Edit"):
            return
        if time.time() - last_notice > 5:
            print("Waiting for RMS login / 2FA completion in the visible browser")
            last_notice = time.time()
        time.sleep(1)

    raise TimeoutException("Timed out waiting for RMS login / 2FA completion")


def maybe_confirm_modal(builder: RMSBudgetBuilder):
    selectors = [
        ".bootbox.modal.show .btn-primary",
        ".bootbox.modal.show .btn-danger",
        ".modal.show .btn-primary",
        ".modal.show .btn-danger",
    ]
    for selector in selectors:
        elements = visible_elements(builder.driver, selector)
        if elements:
            builder.driver.execute_script("arguments[0].click();", elements[0])
            time.sleep(0.7)
            return True
    return False


def create_element_js(builder: RMSBudgetBuilder, category: BudgetCategory, year: int, name: str):
    add_button = builder.driver.find_element(
        By.CSS_SELECTOR,
        f'#year{year} .-delta-budget-line[data-name="{category.value}"] .-delta-add-item',
    )
    builder.driver.execute_script("arguments[0].click();", add_button)
    wait = WebDriverWait(builder.driver, 20)
    input_box = wait.until(EC.element_to_be_clickable((By.ID, "__bootbox_custom_input")))
    input_box.clear()
    input_box.send_keys(name)
    input_box.send_keys(Keys.ENTER)
    accept = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".bootbox.modal.show .btn-primary"))
    )
    builder.driver.execute_script("arguments[0].click();", accept)
    time.sleep(1.0)


def click_year(builder: RMSBudgetBuilder, year: int):
    builder.goto_budget_year(year)
    WebDriverWait(builder.driver, 10).until(
        lambda driver: "active" in driver.find_element(By.ID, f"year{year}").get_attribute("class")
    )
    time.sleep(0.4)


def goto_project_budget(builder: RMSBudgetBuilder, proposal_id: str):
    builder.driver.get(f"https://rms.arc.gov.au/RMS/Proposal/Form/Edit/{proposal_id}")
    wait = WebDriverWait(builder.driver, 30)
    wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                '//button[contains(normalize-space(.), "Project Cost")]',
            )
        )
    )
    project_cost_button = builder.driver.find_element(
        By.XPATH,
        '//button[contains(normalize-space(.), "Project Cost")]',
    )
    builder.driver.execute_script("arguments[0].click();", project_cost_button)
    wait.until(EC.presence_of_element_located((By.ID, "-delta-budget-year-1")))
    time.sleep(1)


def clear_category(builder: RMSBudgetBuilder, year: int, category: str):
    while True:
        selector = (
            f'#year{year} .-delta-budget-line[data-parent="{category}"] '
            ".-delta-delete-item"
        )
        buttons = visible_elements(builder.driver, selector)
        if not buttons:
            return
        builder.driver.execute_script("arguments[0].click();", buttons[-1])
        time.sleep(0.3)
        maybe_confirm_modal(builder)
        time.sleep(1.0)


def clear_year(builder: RMSBudgetBuilder, year: int):
    click_year(builder, year)
    for category in ["Personnel", "Travel", "Other"]:
        clear_category(builder, year, category)
    print(f"Cleared year {year}")


def desired_names_for_payload(payload: dict, category: str):
    return {
        entry["name"]
        for year_data in payload["years"]
        for entry in year_data["entries"]
        if entry["category"] == category
    }


def remove_extra_rows(builder: RMSBudgetBuilder, payload: dict, year: int = 1):
    click_year(builder, year)
    removed = False
    for category in ["Personnel", "Travel", "Other"]:
        desired = desired_names_for_payload(payload, category)
        selector = (
            f'#year{year} .-delta-budget-line[data-parent="{category}"]:not([data-template="true"])'
        )
        while True:
            rows = visible_elements(builder.driver, selector)
            extra_row = None
            extra_name = ""
            for row in reversed(rows):
                name = (row.get_attribute("data-name") or "").strip()
                if name and name not in desired:
                    extra_row = row
                    extra_name = name
                    break
            if extra_row is None:
                break
            delete_button = extra_row.find_element(By.CSS_SELECTOR, ".-delta-delete-item")
            builder.driver.execute_script("arguments[0].click();", delete_button)
            time.sleep(0.3)
            maybe_confirm_modal(builder)
            time.sleep(1.0)
            removed = True
            print(f"Removed extra row from year {year}: {category} | {extra_name}")
    return removed


def to_budget_category(name: str) -> BudgetCategory:
    return BudgetCategory[name.upper().replace(" ", "_")]


def row_selector(year: int, category: str, name: str) -> str:
    return (
        f'#year{year} .-delta-budget-line[data-parent="{category}"]'
        f'[data-name="{name}"]'
    )


def ensure_entry_row(builder: RMSBudgetBuilder, year: int, category: str, name: str):
    selector = row_selector(year, category, name)
    rows = builder.driver.find_elements(By.CSS_SELECTOR, selector)
    if rows:
        return rows[0]

    labeled_rows = visible_elements(
        builder.driver,
        f'#year{year} .-delta-budget-line[data-parent="{category}"]',
    )
    for row in labeled_rows:
        label_elements = row.find_elements(By.CSS_SELECTOR, ".-delta-category-name")
        label = label_elements[0].text.strip() if label_elements else ""
        if label == name:
            return row

    category_rows_selector = (
        f'#year{year} .-delta-budget-line[data-parent="{category}"]:not([data-template="true"])'
    )
    existing_count = len(builder.driver.find_elements(By.CSS_SELECTOR, category_rows_selector))
    add_selector = (
        f'#year{year} .-delta-budget-line[data-name="{category}"] .-delta-add-item'
    )
    add_buttons = builder.driver.find_elements(By.CSS_SELECTOR, add_selector)
    if not add_buttons:
        debug_dir = ROOT / "tmp/web-debug"
        debug_dir.mkdir(parents=True, exist_ok=True)
        (debug_dir / f"missing-add-button-year-{year}.html").write_text(builder.driver.page_source)
        builder.driver.save_screenshot(str(debug_dir / f"missing-add-button-year-{year}.png"))
        raise RuntimeError(f"No add button found for year {year}, category {category}, selector {add_selector}")
    add_button = add_buttons[0]
    builder.driver.execute_script("arguments[0].click();", add_button)
    wait = WebDriverWait(builder.driver, 20)
    input_box = wait.until(EC.visibility_of_element_located((By.ID, "__bootbox_custom_input")))
    input_box.clear()
    input_box.send_keys(name)
    input_box.send_keys(Keys.ENTER)
    accept = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".bootbox.modal.show .btn-primary"))
    )
    builder.driver.execute_script("arguments[0].click();", accept)
    try:
        wait.until(
            lambda driver: len(driver.find_elements(By.CSS_SELECTOR, category_rows_selector))
            > existing_count
        )
    except TimeoutException:
        labeled_rows = visible_elements(
            builder.driver,
            f'#year{year} .-delta-budget-line[data-parent="{category}"]',
        )
        for row in labeled_rows:
            label_elements = row.find_elements(By.CSS_SELECTOR, ".-delta-category-name")
            label = label_elements[0].text.strip() if label_elements else ""
            if label == name:
                return row

        debug_dir = ROOT / "tmp/web-debug"
        debug_dir.mkdir(parents=True, exist_ok=True)
        safe_category = category.lower().replace(" ", "-")
        (debug_dir / f"row-create-timeout-year-{year}-{safe_category}.html").write_text(
            builder.driver.page_source
        )
        builder.driver.save_screenshot(
            str(debug_dir / f"row-create-timeout-year-{year}-{safe_category}.png")
        )
        current_names = [
            element.get_attribute("data-name")
            for element in builder.driver.find_elements(By.CSS_SELECTOR, category_rows_selector)
        ]
        raise TimeoutException(
            f"Timed out creating row {name!r} in {category} year {year}. "
            f"Current rows: {current_names}"
        )
    time.sleep(0.7)
    return builder.driver.find_elements(By.CSS_SELECTOR, category_rows_selector)[-1]


def set_input_value(builder: RMSBudgetBuilder, element, value: int):
    builder.driver.execute_script(
        """
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
        arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
        """,
        element,
        str(value),
    )


def normalized_inputs(row):
    values = [item.get_attribute("value") for item in row.find_elements(By.CSS_SELECTOR, "input.-delta-input")]
    normalized = []
    for value in values:
        value = (value or "").replace(",", "").strip()
        normalized.append(value or "0")
    return normalized


def desired_inputs(entry: dict, input_count: int):
    values = [str(entry["arc"]), str(entry["admin"])]
    if input_count >= 3:
        values.append(str(entry["inkind"]))
    return values


def fill_entry(builder: RMSBudgetBuilder, year: int, entry: dict):
    row = ensure_entry_row(builder, year, entry["category"], entry["name"])
    inputs = row.find_elements(By.CSS_SELECTOR, "input.-delta-input")
    if len(inputs) < 2:
        raise RuntimeError(f"Unexpected inputs for {entry['name']}: {len(inputs)}")

    current = normalized_inputs(row)
    desired = desired_inputs(entry, len(inputs))
    changed = False

    for input_element, current_value, desired_value in zip(inputs, current, desired):
        if current_value == desired_value:
            continue
        set_input_value(builder, input_element, int(desired_value))
        changed = True

    return changed


def fill_year(builder: RMSBudgetBuilder, year_data: dict):
    year = int(year_data["year"])
    click_year(builder, year)
    changed = False
    for entry in year_data["entries"]:
        entry_changed = fill_entry(builder, year, entry)
        if entry_changed:
            changed = True
            print(
                f"Updated year {year}: {entry['category']} | {entry['name']} | "
                f"ARC {entry['arc']} | Admin {entry['admin']} | In-kind {entry['inkind']}"
            )
        else:
            print(
                f"Unchanged year {year}: {entry['category']} | {entry['name']}"
            )
    return changed


def print_totals(builder: RMSBudgetBuilder, year: int):
    click_year(builder, year)
    selector = (
        f'#year{year} .-delta-budget-line[data-parent=""][data-name="Total"] '
        ".-delta-value"
    )
    values = [element.get_attribute("data-amount") for element in visible_elements(builder.driver, selector)]
    values = [value for value in values if value]
    print(f"Year {year} totals after save: {values}")


def scrape_visible_rows(builder: RMSBudgetBuilder, year: int):
    click_year(builder, year)
    rows = visible_elements(
        builder.driver,
        f"#year{year} .-delta-budget-line:not([data-template='true'])",
    )
    output = []
    for row in rows:
        name = row.get_attribute("data-name") or ""
        parent = row.get_attribute("data-parent") or ""
        inputs = row.find_elements(By.CSS_SELECTOR, "input.-delta-input")
        input_values = [item.get_attribute("value") for item in inputs]
        if name:
            output.append({"parent": parent, "name": name, "inputs": input_values})
    return output


def main():
    args = parse_args()
    builder = RMSBudgetBuilder()
    options = ChromeOptions()
    options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    options.add_argument("--window-size=1440,1000")
    options.add_argument("--no-sandbox")
    builder.driver = webdriver.Chrome(
        service=ChromeService(executable_path=str(resolve_matching_chromedriver())),
        options=options,
    )
    builder.create_element = lambda category, year, name: create_element_js(
        builder, category, year, name
    )
    login_interactive(builder)
    goto_project_budget(builder, PAYLOAD["proposal_id"])

    if args.full_reset:
        for year in range(1, 6):
            clear_year(builder, year)

        builder.save_budget()
        time.sleep(3)
        goto_project_budget(builder, PAYLOAD["proposal_id"])

    any_changed = False
    for year_data in PAYLOAD["years"]:
        any_changed = fill_year(builder, year_data) or any_changed

    if any_changed or args.full_reset:
        builder.save_budget()
        time.sleep(3)
        goto_project_budget(builder, PAYLOAD["proposal_id"])

    if not args.no_prune:
        for cleanup_round in range(1, 8):
            removed_any = False
            for year in [1, 2, 3]:
                removed_any = remove_extra_rows(builder, PAYLOAD, year) or removed_any
            if not removed_any:
                break
            print(f"Completed cleanup round {cleanup_round}")
            builder.save_budget()
            time.sleep(3)
            goto_project_budget(builder, PAYLOAD["proposal_id"])

    if not args.no_prune:
        builder.save_budget()
        time.sleep(3)
        goto_project_budget(builder, PAYLOAD["proposal_id"])

    for year in [1, 2, 3]:
        print_totals(builder, year)

    verify = {"proposal_id": PAYLOAD["proposal_id"], "years": []}
    for year in [1, 2, 3]:
        verify["years"].append({"year": year, "rows": scrape_visible_rows(builder, year)})
    verify_path = ROOT / "tmp" / "rms_verify_current.json"
    verify_path.write_text(json.dumps(verify, indent=2))
    print(f"Wrote verification to {verify_path}")

    screenshot_path = ROOT / "tmp/web-debug/rms-budget-after-latest-upload.png"
    builder.driver.save_screenshot(str(screenshot_path))
    print(f"Saved latest budget to RMS and screenshot to {screenshot_path}")


if __name__ == "__main__":
    main()

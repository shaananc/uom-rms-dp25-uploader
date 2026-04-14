from __future__ import annotations

import json
import logging
import subprocess
import time
from enum import Enum
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

logger = logging.getLogger(__name__)


class BudgetCategory(Enum):
    PERSONNEL = "Personnel"
    TEACHING_RELIEF = "Teaching Relief"
    TRAVEL = "Travel"
    FIELD_RESEARCH = "Field Research"
    EQUIPMENT = "Equipment"
    MAINTENANCE = "Maintenance"
    OTHER = "Other"
    TOTAL = "Total"


class RMSBudgetBuilder:
    def __init__(
        self,
        *,
        root: Path | None = None,
        proposal_id: str | None = None,
        driver: webdriver.Chrome | None = None,
        chrome_binary: str = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        chromedriver_path: str | None = None,
    ) -> None:
        self.root = root or Path(__file__).resolve().parent
        self.proposal_id = proposal_id
        self.driver = driver
        self.chrome_binary = chrome_binary
        self.chromedriver_path = chromedriver_path

    def __del__(self) -> None:
        try:
            if self.driver is not None:
                self.driver.quit()
        except Exception:
            pass

    @staticmethod
    def visible_elements(driver, selector: str):
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        return [element for element in elements if element.is_displayed()]

    @staticmethod
    def canonical_name(name: str) -> str:
        return str(name).splitlines()[0].strip()

    def chrome_options(self) -> ChromeOptions:
        options = ChromeOptions()
        options.binary_location = self.chrome_binary
        options.add_argument("--window-size=1440,1000")
        options.add_argument("--no-sandbox")
        return options

    def resolve_matching_chromedriver(self) -> str:
        chrome_binary = Path(self.chrome_binary)
        version_output = subprocess.check_output(
            [str(chrome_binary), "--version"], text=True
        ).strip()
        chrome_version = version_output.removeprefix("Google Chrome ").strip()
        cached_driver = (
            Path.home()
            / ".cache/selenium/chromedriver/mac-arm64"
            / chrome_version
            / "chromedriver"
        )
        if cached_driver.exists():
            return str(cached_driver)
        raise FileNotFoundError(
            f"No matching cached ChromeDriver for Chrome {chrome_version}"
        )

    def setup(self) -> None:
        options = self.chrome_options()
        driver_path = self.chromedriver_path or self.resolve_matching_chromedriver()
        service = ChromeService(executable_path=driver_path)
        self.driver = webdriver.Chrome(service=service, options=options)

    def read_credentials(self, path: Path | None = None) -> dict[str, str]:
        credentials_path = path or (self.root / "rmscreds.txt")
        lines = credentials_path.read_text().splitlines()
        return {"username": lines[0].strip(), "password": lines[1].strip()}

    def login(self, credentials_path: Path | None = None) -> None:
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        credentials = self.read_credentials(credentials_path)
        self.driver.get("https://rms.arc.gov.au/RMS/ActionCentre")
        self.driver.set_window_size(550, 691)

        if self.driver.find_elements(By.LINK_TEXT, "Edit"):
            return

        if self.driver.find_elements(By.ID, "emailAddress"):
            self.driver.find_element(By.ID, "emailAddress").send_keys(
                credentials["username"]
            )
            password = self.driver.find_element(By.ID, "password")
            password.click()
            password.send_keys(credentials["password"])
            try:
                self.driver.find_element(By.ID, "login").click()
            except Exception:
                self.driver.execute_script(
                    "arguments[0].click();", self.driver.find_element(By.ID, "login")
                )

        deadline = time.time() + 10 * 60
        last_notice = 0.0
        while time.time() < deadline:
            if self.driver.find_elements(By.LINK_TEXT, "Edit"):
                return
            if time.time() - last_notice > 5:
                logger.info(
                    "Waiting for RMS login / 2FA completion in the visible browser"
                )
                last_notice = time.time()
            time.sleep(1)

        raise TimeoutException("Timed out waiting for RMS login / 2FA completion")

    def goto_budget(self, proposal_id: str | None = None) -> None:
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        target_proposal = proposal_id or self.proposal_id
        if not target_proposal:
            raise ValueError("proposal_id must be provided")

        self.driver.get(f"https://rms.arc.gov.au/RMS/Proposal/Form/Edit/{target_proposal}")
        wait = WebDriverWait(self.driver, 30)
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//button[contains(normalize-space(.), "Project Cost")]')
            )
        )
        project_cost_button = self.driver.find_element(
            By.XPATH, '//button[contains(normalize-space(.), "Project Cost")]'
        )
        self.driver.execute_script("arguments[0].click();", project_cost_button)
        wait.until(EC.presence_of_element_located((By.ID, "-delta-budget-year-1")))
        time.sleep(1)

    def goto_budget_year(self, year: int) -> None:
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        year_button = self.driver.find_element(By.ID, f"-delta-budget-year-{year}")
        self.driver.execute_script("arguments[0].click();", year_button)
        WebDriverWait(self.driver, 10).until(
            lambda driver: "active"
            in driver.find_element(By.ID, f"year{year}").get_attribute("class")
        )
        time.sleep(0.4)

    def maybe_confirm_modal(self) -> bool:
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        selectors = [
            ".bootbox.modal.show .btn-primary",
            ".bootbox.modal.show .btn-danger",
            ".modal.show .btn-primary",
            ".modal.show .btn-danger",
        ]
        for selector in selectors:
            elements = self.visible_elements(self.driver, selector)
            if elements:
                self.driver.execute_script("arguments[0].click();", elements[0])
                time.sleep(0.7)
                return True
        return False

    def create_element(self, category: BudgetCategory, year: int, name: str) -> None:
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        add_button = self.driver.find_element(
            By.CSS_SELECTOR,
            f'#year{year} .-delta-budget-line[data-name="{category.value}"] .-delta-add-item',
        )
        self.driver.execute_script("arguments[0].click();", add_button)
        wait = WebDriverWait(self.driver, 20)
        input_box = wait.until(EC.element_to_be_clickable((By.ID, "__bootbox_custom_input")))
        input_box.clear()
        input_box.send_keys(name)
        input_box.send_keys("\n")
        accept = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".bootbox.modal.show .btn-primary"))
        )
        self.driver.execute_script("arguments[0].click();", accept)
        time.sleep(1.0)

    def row_selector(self, year: int, category: str, name: str) -> str:
        return (
            f'#year{year} .-delta-budget-line[data-parent="{category}"]'
            f'[data-name="{name}"]'
        )

    def clear_category(self, year: int, category: str) -> None:
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        while True:
            selector = (
                f'#year{year} .-delta-budget-line[data-parent="{category}"] .-delta-delete-item'
            )
            buttons = self.visible_elements(self.driver, selector)
            if not buttons:
                return
            self.driver.execute_script("arguments[0].click();", buttons[-1])
            time.sleep(0.3)
            self.maybe_confirm_modal()
            time.sleep(1.0)

    def clear_year(self, year: int) -> None:
        self.goto_budget_year(year)
        for category in ["Personnel", "Travel", "Other"]:
            self.clear_category(year, category)
        logger.info("Cleared year %s", year)

    def desired_names_for_payload(self, payload: dict, category: str) -> set[str]:
        return {
            self.canonical_name(entry["name"])
            for year_data in payload["years"]
            for entry in year_data["entries"]
            if entry["category"] == category
        }

    def remove_extra_rows(self, payload: dict, year: int) -> bool:
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        self.goto_budget_year(year)
        removed = False
        for category in ["Personnel", "Travel", "Other"]:
            desired = self.desired_names_for_payload(payload, category)
            selector = (
                f'#year{year} .-delta-budget-line[data-parent="{category}"]:not([data-template="true"])'
            )
            while True:
                rows = self.visible_elements(self.driver, selector)
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
                self.driver.execute_script("arguments[0].click();", delete_button)
                time.sleep(0.3)
                self.maybe_confirm_modal()
                time.sleep(1.0)
                removed = True
                logger.info(
                    "Removed extra row from year %s: %s | %s",
                    year,
                    category,
                    extra_name,
                )
        return removed

    def ensure_entry_row(self, year: int, category: str, name: str):
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        selector = self.row_selector(year, category, name)
        rows = self.driver.find_elements(By.CSS_SELECTOR, selector)
        if rows:
            return rows[0]

        labeled_rows = self.visible_elements(
            self.driver,
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
        existing_count = len(self.driver.find_elements(By.CSS_SELECTOR, category_rows_selector))
        self.create_element(BudgetCategory[category.upper().replace(" ", "_")], year, name)
        WebDriverWait(self.driver, 20).until(
            lambda driver: len(driver.find_elements(By.CSS_SELECTOR, category_rows_selector))
            > existing_count
        )
        time.sleep(0.7)
        return self.driver.find_elements(By.CSS_SELECTOR, category_rows_selector)[-1]

    def set_input_value(self, element, value: int) -> None:
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        self.driver.execute_script(
            """
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
            arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
            """,
            element,
            str(value),
        )

    def normalized_inputs(self, row) -> list[str]:
        values = [
            item.get_attribute("value")
            for item in row.find_elements(By.CSS_SELECTOR, "input.-delta-input")
        ]
        normalized: list[str] = []
        for value in values:
            value = (value or "").replace(",", "").strip()
            normalized.append(value or "0")
        return normalized

    def desired_inputs(self, arc_cash: int, admin_cash: int, admin_inkind: int, input_count: int) -> list[str]:
        values = [str(arc_cash), str(admin_cash)]
        if input_count >= 3 and admin_inkind != -1:
            values.append(str(admin_inkind))
        return values

    def input_category(
        self,
        year: int,
        category: BudgetCategory,
        name: str,
        arc_cash: int,
        admin_cash: int,
        admin_inkind: int,
    ) -> bool:
        row = self.ensure_entry_row(year, category.value, name.strip())
        inputs = row.find_elements(By.CSS_SELECTOR, "input.-delta-input")
        desired = self.desired_inputs(arc_cash, admin_cash, admin_inkind, len(inputs))
        current = self.normalized_inputs(row)

        changed = False
        for input_element, current_value, desired_value in zip(inputs, current, desired):
            if current_value == desired_value:
                continue
            self.set_input_value(input_element, int(desired_value))
            changed = True
        return changed

    def save_budget(self) -> None:
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        save_button = self.driver.find_element(By.ID, "-delta-form-save")
        self.driver.execute_script("arguments[0].click();", save_button)
        WebDriverWait(self.driver, 30).until(
            lambda current: current.find_element(By.ID, "-delta-form-save").is_enabled()
        )
        time.sleep(0.5)

    def scrape_visible_rows(self, year: int) -> list[dict[str, object]]:
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        self.goto_budget_year(year)
        rows = self.visible_elements(
            self.driver,
            f"#year{year} .-delta-budget-line:not([data-template='true'])",
        )
        output: list[dict[str, object]] = []
        for row in rows:
            name = row.get_attribute("data-name") or ""
            parent = row.get_attribute("data-parent") or ""
            inputs = row.find_elements(By.CSS_SELECTOR, "input.-delta-input")
            input_values = [item.get_attribute("value") for item in inputs]
            if name:
                output.append({"parent": parent, "name": name, "inputs": input_values})
        return output

    def print_totals(self, year: int) -> None:
        if self.driver is None:
            raise RuntimeError("Driver is not initialized")

        self.goto_budget_year(year)
        selector = (
            f'#year{year} .-delta-budget-line[data-parent=""][data-name="Total"] .-delta-value'
        )
        values = [
            element.get_attribute("data-amount")
            for element in self.visible_elements(self.driver, selector)
        ]
        values = [value for value in values if value]
        logger.info("Year %s totals after save: %s", year, values)

    def sync_entry(self, year: int, entry: dict) -> bool:
        return self.input_category(
            year,
            BudgetCategory[entry["category"].upper().replace(" ", "_")],
            self.canonical_name(entry["name"]),
            int(entry["arc"]),
            int(entry["admin"]),
            int(entry["inkind"]),
        )

    def sync_year(self, year_data: dict) -> bool:
        year = int(year_data["year"])
        self.goto_budget_year(year)
        changed = False
        for entry in year_data["entries"]:
            entry_changed = self.sync_entry(year, entry)
            if entry_changed:
                changed = True
                logger.info(
                    "Updated year %s: %s | %s | ARC %s | Admin %s | In-kind %s",
                    year,
                    entry["category"],
                    entry["name"],
                    entry["arc"],
                    entry["admin"],
                    entry["inkind"],
                )
            else:
                logger.info(
                    "Unchanged year %s: %s | %s",
                    year,
                    entry["category"],
                    entry["name"],
                )
        return changed

    def sync_payload(
        self,
        payload: dict,
        *,
        full_reset: bool = False,
        prune: bool = True,
        years: list[int] | None = None,
    ) -> None:
        target_years = years or [int(year_data["year"]) for year_data in payload["years"]]

        if full_reset:
            for year in range(1, 6):
                self.clear_year(year)
            self.save_budget()
            self.goto_budget(payload["proposal_id"])

        any_changed = False
        for year_data in payload["years"]:
            any_changed = self.sync_year(year_data) or any_changed

        if any_changed or full_reset:
            self.save_budget()
            self.goto_budget(payload["proposal_id"])

        if prune:
            for cleanup_round in range(1, 8):
                removed_any = False
                for year in target_years:
                    removed_any = self.remove_extra_rows(payload, year) or removed_any
                if not removed_any:
                    break
                logger.info("Completed cleanup round %s", cleanup_round)
                self.save_budget()
                self.goto_budget(payload["proposal_id"])
            self.save_budget()
            self.goto_budget(payload["proposal_id"])

    def write_verification(self, path: Path, years: list[int]) -> None:
        if not self.proposal_id:
            raise ValueError("proposal_id must be set to write verification")

        verify = {"proposal_id": self.proposal_id, "years": []}
        for year in years:
            verify["years"].append({"year": year, "rows": self.scrape_visible_rows(year)})
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(verify, indent=2))

from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path

from rich.logging import RichHandler

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from RMSBudgetUploader import RMSBudgetBuilder

logging.basicConfig(level=logging.INFO, format="%(message)s", handlers=[RichHandler()])


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Sync the current RMS budget payload to RMS."
    )
    parser.add_argument(
        "--payload",
        type=Path,
        default=ROOT / "tmp" / "rms_current_payload.json",
        help="Path to the generated RMS payload JSON.",
    )
    parser.add_argument(
        "--credentials",
        type=Path,
        default=ROOT / "rmscreds.txt",
        help="Two-line credentials file: email then password.",
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


def load_payload(path: Path) -> dict:
    payload = json.loads(path.read_text())
    for year_data in payload["years"]:
        for entry in year_data["entries"]:
            entry["name"] = RMSBudgetBuilder.canonical_name(entry["name"])
    return payload


def main() -> None:
    args = parse_args()
    payload = load_payload(args.payload)

    builder = RMSBudgetBuilder(root=ROOT, proposal_id=payload["proposal_id"])
    builder.setup()
    builder.login(args.credentials)
    builder.goto_budget()
    builder.sync_payload(payload, full_reset=args.full_reset, prune=not args.no_prune)

    for year in [1, 2, 3]:
        builder.print_totals(year)

    verify_path = ROOT / "tmp" / "rms_verify_current.json"
    builder.write_verification(verify_path, [1, 2, 3])
    print(f"Wrote verification to {verify_path}")

    screenshot_path = ROOT / "tmp" / "web-debug" / "rms-budget-after-latest-upload.png"
    screenshot_path.parent.mkdir(parents=True, exist_ok=True)
    builder.driver.save_screenshot(str(screenshot_path))
    print(f"Saved latest budget to RMS and screenshot to {screenshot_path}")


if __name__ == "__main__":
    main()

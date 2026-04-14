from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"main": MAIN_NS}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Update in-kind contribution cells in the reconstructed ARC budget workbook."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=Path("output/spreadsheet/DP27-Budget-Reconstructed.xlsx"),
        help="Workbook to modify in place.",
    )
    return parser.parse_args()


def set_numeric_cell(root: ET.Element, cell_ref: str, value: int) -> None:
    cell = root.find(f".//main:c[@r='{cell_ref}']", NS)
    if cell is None:
        raise ValueError(f"Could not find cell {cell_ref}")

    cell.set("t", "n")
    value_node = cell.find("main:v", NS)
    if value_node is None:
        value_node = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
    value_node.text = str(value)


def set_inline_string_cell(root: ET.Element, cell_ref: str, value: str) -> None:
    cell = root.find(f".//main:c[@r='{cell_ref}']", NS)
    if cell is None:
        raise ValueError(f"Could not find cell {cell_ref}")

    cell.set("t", "inlineStr")
    for child in list(cell):
        if child.tag != f"{{{MAIN_NS}}}is":
            cell.remove(child)

    inline = cell.find("main:is", NS)
    if inline is None:
        inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")

    for child in list(inline):
        inline.remove(child)

    text = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
    text.text = value


def update_workbook(workbook_path: Path) -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        with zipfile.ZipFile(workbook_path) as source_zip:
            source_zip.extractall(tmpdir_path)

        sheet_path = tmpdir_path / "xl" / "worksheets" / "sheet2.xml"
        root = ET.parse(sheet_path).getroot()

        # Xingliang is already budgeted at Level D.4 in older repo artifacts.
        # For 2027 and 2028 the in-kind entry is only the UoM top-up above the
        # fully costed ARC Future Fellowship Step 2 total. The fellowship is
        # assumed to end in March 2029, so the 2029 value is prorated for
        # three months of fellowship coverage and nine months at the full UoM
        # D.4 rate.
        set_inline_string_cell(
            root,
            "A15",
            "CI Xingliang Yuan, Level D.4 at 0.2 FTE (plus on-costs)",
        )
        set_numeric_cell(root, "E15", 8339)
        set_numeric_cell(root, "L15", 8339)
        set_numeric_cell(root, "S15", 43989)

        # Shaanan starts the project at Level C.3. The in-kind entry is only
        # the UoM top-up above the fully costed ARC DECRA total.
        set_numeric_cell(root, "E16", 18751)
        set_numeric_cell(root, "L16", 20038)
        set_numeric_cell(root, "S16", 21327)

        ET.register_namespace("", MAIN_NS)
        ET.ElementTree(root).write(sheet_path, encoding="utf-8", xml_declaration=False)

        rebuilt_path = tmpdir_path / "rebuilt.xlsx"
        with zipfile.ZipFile(rebuilt_path, "w", compression=zipfile.ZIP_DEFLATED) as rebuilt:
            for path in sorted(tmpdir_path.rglob("*")):
                if path == rebuilt_path or path.is_dir():
                    continue
                rebuilt.write(path, path.relative_to(tmpdir_path))

        shutil.move(rebuilt_path, workbook_path)


def main() -> None:
    args = parse_args()
    update_workbook(args.workbook)


if __name__ == "__main__":
    main()

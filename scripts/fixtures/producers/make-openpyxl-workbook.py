#!/usr/bin/env python3
"""Write the F22 openpyxl producer fixture.

Called by scripts/fixtures/groups/producers.mjs; the module normalises the
ZIP container and document timestamps afterwards. openpyxl does not calculate
formulas, so the formula cell is written without a cached value on purpose:
that is the producer behaviour the open path must accept (file-compat §6.1).
"""

import sys
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font


def main(output_path: str) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Producer"

    sheet["A1"] = "Producer"
    sheet["A1"].font = Font(bold=True)
    sheet["B1"] = "openpyxl"

    sheet["A3"] = "Header"
    sheet["B3"] = "Value"
    sheet["A4"] = "Amount"
    sheet["B4"] = 1234.5
    sheet["B4"].number_format = "#,##0.00"
    sheet["A5"] = "Doubled"
    sheet["B5"] = "=B4*2"
    sheet["A6"] = "Accent"
    sheet["B6"] = "café ☃"
    sheet["A7"] = "Date"
    sheet["B7"] = datetime(2026, 11, 19)
    sheet["B7"].number_format = "yyyy-mm-dd"

    second = workbook.create_sheet("Second")
    second["A1"] = "A second sheet"

    workbook.save(output_path)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("usage: make-openpyxl-workbook.py <output.xlsx>")
    main(sys.argv[1])

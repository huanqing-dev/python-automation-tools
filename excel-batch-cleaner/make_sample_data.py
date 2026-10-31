"""
Generate 3 intentionally messy Excel files for demo purposes.
sales_mar.xlsx has 2 sheets to demo multi-sheet merging.
Run this first, then run excel_cleaner.py to see the cleanup in action.
"""

from pathlib import Path

import pandas as pd

COLUMNS = [" Customer ", " Product", "Qty ", " Date "]  # messy headers too

# Messy on purpose: extra spaces, empty rows, duplicate rows.
BOOKS = {
    "sales_jan.xlsx": [
        ("Sheet1", [
            [" Alice ", "Widget A", 3, "2026-01-05"],
            ["Bob", "Widget B", 1, "2026-01-06"],
            ["Alice", "Widget A", 3, "2026-01-05"],  # dupe after stripping
            [None, None, None, None],                # empty row
            [" carol", "Widget C ", 2, "2026-01-07"],
        ]),
    ],
    "sales_feb.xlsx": [
        ("Sheet1", [
            ["Dave", "Widget A", 5, "2026-02-02"],
            ["Eve ", " Widget B", 2, "2026-02-03"],
            [None, None, None, None],                # empty row
            ["dave", "Widget A", 5, "2026-02-02"],   # NOT a dupe (case differs)
        ]),
    ],
    "sales_mar.xlsx": [
        ("orders", [
            ["Frank", "Widget C", 4, "2026-03-01"],
            [" Grace", "Widget A", 1, "2026-03-02"],
            ["Frank", "Widget C", 4, "2026-03-01"],  # exact duplicate
        ]),
        ("returns", [
            ["Hank", "Widget B", 1, "2026-03-05"],
            ["Ivy ", "Widget A", 2, "2026-03-06"],
        ]),
    ],
}


def main() -> None:
    out_dir = Path("sample_data")
    out_dir.mkdir(exist_ok=True)
    for name, sheets in BOOKS.items():
        with pd.ExcelWriter(out_dir / name) as writer:
            for sheet_name, rows in sheets:
                pd.DataFrame(rows, columns=COLUMNS).to_excel(
                    writer, sheet_name=sheet_name, index=False)
        print(f"created {out_dir / name} ({len(sheets)} sheet(s))")


if __name__ == "__main__":
    main()

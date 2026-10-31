#!/usr/bin/env python3
"""Turn PDF tables into a clean, formatted Excel workbook.

Why this exists: copying a table out of a PDF by hand is slow and lossy. Most
"PDF to Excel" tools dump everything into one sheet, keep the repeated header
rows, and turn every number into text. This tool does the opposite.

  * Detects ruled *and* border-less tables (three extraction strategies, best
    non-overlapping candidates win)
  * Rejoins a table that continues across pages, and drops the header that the
    PDF repeats on each page
  * Drops page footers such as "Page 2 of 3"
  * Coerces numbers ("1,234.50", "(480.00)", "6%") into real Excel numbers with
    a matching format - but leaves a column as text when it mixes currencies
  * Writes one sheet per table: bold frozen header, auto-filter, sized columns
  * Prints a report of what was kept, merged and discarded

Usage
    python pdf_tables.py                            # every PDF in ./sample_data
    python pdf_tables.py invoice.pdf                # a single file
    python pdf_tables.py "reports/*.pdf" -o out.xlsx

Requires: pdfplumber, openpyxl
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from pathlib import Path
from types import SimpleNamespace

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Three ways of seeing a table, most reliable first. "lines" needs the PDF to
# draw the grid; the text strategies fall back to column positions.
STRATEGIES = (
    {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
    {"vertical_strategy": "lines", "horizontal_strategy": "text"},
    {"vertical_strategy": "text", "horizontal_strategy": "text", "text_x_tolerance": 3},
)

CURRENCY = "¥$€£"
FOOTER_RE = re.compile(
    r"^(page\s*\d+\s*(of|/)\s*\d+|\d+\s*/\s*\d+|continued|confidential|"
    r"internal report.*|第\s*\d+\s*页.*|.*do not distribute)$",
    re.IGNORECASE,
)
DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y", "%Y/%m/%d")


# --------------------------------------------------------------------------- #
# extraction
# --------------------------------------------------------------------------- #
def _overlap(a, b) -> float:
    """Share of the smaller bounding box that is covered by the other one."""
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    ix = max(0.0, min(ax1, bx1) - max(ax0, bx0))
    iy = max(0.0, min(ay1, by1) - max(ay0, by0))
    inter = ix * iy
    if inter <= 0:
        return 0.0
    small = min((ax1 - ax0) * (ay1 - ay0), (bx1 - bx0) * (by1 - by0))
    return inter / small if small else 0.0


def _clean_cell(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ")).strip()


def _residual_bands(bbox, taken) -> list[tuple[float, float]]:
    """Vertical bands of `bbox` that no accepted table covers.

    A loose text-based strategy often swallows a ruled table *and* the
    border-less table underneath it. Rather than dropping the whole candidate we
    keep the part that is still free, so the second table is not lost.
    """
    x0, y0, x1, y1 = bbox
    bands = [(y0, y1)]
    for t in taken:
        if t[2] <= x0 or t[0] >= x1:  # no horizontal overlap
            continue
        cut0, cut1 = t[1], t[3]
        nxt = []
        for b0, b1 in bands:
            if cut1 <= b0 or cut0 >= b1:
                nxt.append((b0, b1))
                continue
            if cut0 > b0:
                nxt.append((b0, cut0))
            if cut1 < b1:
                nxt.append((cut1, b1))
        bands = nxt
    return [(b0, b1) for b0, b1 in bands if b1 - b0 > 6]


def _harvest(target, settings) -> list[SimpleNamespace]:
    out = []
    try:
        found = target.find_tables(table_settings=settings)
    except Exception:  # a broken page must not kill the run
        return out
    for table in found:
        rows = [[_clean_cell(c) for c in row] for row in table.extract()]
        rows = [r for r in rows if any(r)]
        if len(rows) < 2 or max(len(r) for r in rows) < 2:
            continue
        out.append(SimpleNamespace(bbox=table.bbox, rows=rows))
    return out


def page_tables(page) -> list[SimpleNamespace]:
    """All distinct tables on one page, in reading order."""
    accepted: list[SimpleNamespace] = []
    for settings in STRATEGIES:
        for table in _harvest(page, settings):
            if not any(_overlap(table.bbox, other.bbox) > 0.5 for other in accepted):
                accepted.append(table)  # clean table, better strategy won
                continue
            # Partly covered: re-detect inside the free bands only.
            for band in _residual_bands(table.bbox, [a.bbox for a in accepted]):
                crop = page.crop((table.bbox[0], band[0], table.bbox[2], band[1]))
                for extra in _harvest(crop, settings):
                    bbox = (
                        extra.bbox[0], band[0] + extra.bbox[1],
                        extra.bbox[2], band[0] + extra.bbox[3],
                    )
                    if any(_overlap(bbox, a.bbox) > 0.5 for a in accepted):
                        continue
                    accepted.append(SimpleNamespace(bbox=bbox, rows=extra.rows))
    accepted.sort(key=lambda t: (round(t.bbox[1], 1), round(t.bbox[0], 1)))
    return accepted


def _is_footer(row: list[str]) -> bool:
    cells = [c for c in row if c]
    return len(cells) == 1 and bool(FOOTER_RE.match(cells[0]))


def _same_header(a: list[str], b: list[str]) -> bool:
    norm = lambda r: [c.lower().replace(" ", "") for c in r if c]
    return bool(a) and norm(a) == norm(b)


def collect(pdf_path: Path) -> list[dict]:
    """Tables of one PDF, with multi-page continuations already stitched."""
    tables: list[dict] = []
    dropped_footers = 0
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            for index, table in enumerate(page_tables(page)):
                rows = []
                for row in table.rows:
                    if _is_footer(row):
                        dropped_footers += 1
                        continue
                    rows.append(row)
                if len(rows) < 2:
                    continue
                header, body = rows[0], rows[1:]
                body = [r for r in body if not _same_header(r, header)]

                previous = tables[-1] if tables else None
                continues = (
                    previous is not None
                    and index == 0
                    and previous["page"] == page_no - 1
                    and _same_header(previous["header"], header)
                )
                if continues:
                    previous["rows"].extend(body)
                    previous["pages"].append(page_no)
                else:
                    tables.append({
                        "header": header,
                        "rows": list(body),
                        "page": page_no,
                        "pages": [page_no],
                    })
    return tables, dropped_footers


# --------------------------------------------------------------------------- #
# typing
# --------------------------------------------------------------------------- #
def _strip_wrapper(text: str):
    """-> (number_or_None, currency_symbol, is_percent, is_negative_wrapped)"""
    raw = text.strip()
    if not raw:
        return None, "", False, False
    negative = raw.startswith("(") and raw.endswith(")")
    if negative:
        raw = raw[1:-1].strip()
    percent = raw.endswith("%")
    if percent:
        raw = raw[:-1].strip()
    symbol = ""
    while raw and raw[0] in CURRENCY:
        symbol += raw[0]
        raw = raw[1:].strip()
    while raw and raw[-1] in CURRENCY:
        symbol += raw[-1]
        raw = raw[:-1].strip()
    candidate = raw.replace(",", "").replace("\u00a0", "").replace(" ", "")
    if not candidate:
        return None, symbol, percent, negative
    try:
        return float(candidate), symbol, percent, negative
    except ValueError:
        return None, symbol, percent, negative


def _as_date(text: str):
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(text.strip(), fmt).date()
        except ValueError:
            continue
    return None


def type_column(values: list[str]) -> dict:
    """Decide how a whole column should be written, from all of its cells."""
    filled = [v for v in values if v]
    if not filled:
        return {"kind": "text"}

    parsed = [_strip_wrapper(v) for v in filled]
    numeric_share = sum(p[0] is not None for p in parsed) / len(filled)

    if numeric_share >= 0.8:
        percents = [p[2] for p in parsed]
        if all(percents):
            cells = [
                (None if n is None else round((-n if neg else n) * 0.01, 10), "0.00%")
                for n, _, _, neg in parsed
            ]
            return {"kind": "number", "cells": cells}

        integral = all(
            n is not None and abs(n - round(n)) < 1e-9 for n, _, _, _ in parsed
        )
        plain_fmt = "#,##0" if integral else "#,##0.00"
        cells = []
        for n, symbol, _, neg in parsed:
            if n is None:
                cells.append((None, ""))
                continue
            value = round(-n if neg else n, 10)
            # Per-cell format keeps a mixed-currency column numeric *and* readable.
            fmt = f'"{symbol}"#,##0.00' if symbol else plain_fmt
            cells.append((value, fmt))
        return {"kind": "number", "cells": cells}

    dates = [_as_date(v) for v in filled]
    if sum(d is not None for d in dates) / len(filled) >= 0.8:
        return {"kind": "date", "cells": [(d, "yyyy-mm-dd") for d in dates]}

    return {"kind": "text"}


def build_matrix(table: dict) -> tuple[list[list], list[list[str]]]:
    """Header + typed body -> (matrix, per-cell number formats)."""
    width = max([len(table["header"])] + [len(r) for r in table["rows"]])
    header = table["header"] + [""] * (width - len(table["header"]))
    body = [r + [""] * (width - len(r)) for r in table["rows"]]
    formats = [[""] * width for _ in range(len(body))]

    for col in range(width):
        spec = type_column([row[col] for row in body])
        if spec["kind"] == "text":
            continue
        typed = iter(spec["cells"])
        for row_index, row in enumerate(body):
            if not row[col]:
                continue
            value, fmt = next(typed)
            row[col] = value
            formats[row_index][col] = fmt
    return [header] + body, formats


# --------------------------------------------------------------------------- #
# workbook
# --------------------------------------------------------------------------- #
HEADER_FILL = PatternFill("solid", fgColor="EEF2F7")


def write_sheet(ws, matrix: list[list], formats: list[list[str]]) -> None:
    for row in matrix:
        ws.append(row)
    for row_index, row_formats in enumerate(formats, start=2):
        for col, fmt in enumerate(row_formats, start=1):
            if fmt:
                ws.cell(row=row_index, column=col).number_format = fmt
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=False)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in range(1, len(matrix[0]) + 1):
        widest = max(
            (len(str(ws.cell(row=r, column=col).value or "")) for r in range(1, len(matrix) + 1)),
            default=8,
        )
        ws.column_dimensions[get_column_letter(col)].width = min(max(10, widest + 2), 60)


def safe_title(name: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]\*\?/\\:]", "-", name).strip() or "Sheet"
    base = base[:31]
    title, n = base, 2
    while title in used:
        suffix = f" ({n})"
        title = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(title)
    return title


def run(inputs: list[Path], out_path: Path) -> int:
    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    total_rows = total_footers = 0

    for pdf_path in inputs:
        tables, footers = collect(pdf_path)
        total_footers += footers
        if not tables:
            print(f"  ! {pdf_path.name}: no tables found")
            continue
        print(f"  {pdf_path.name}: {len(tables)} table(s), {footers} footer row(s) dropped")
        for i, table in enumerate(tables, start=1):
            matrix, formats = build_matrix(table)
            total_rows += len(matrix) - 1
            label = pdf_path.stem if len(tables) == 1 else f"{pdf_path.stem}-{i}"
            ws = wb.create_sheet(safe_title(label, used))
            write_sheet(ws, matrix, formats)
            pages = ",".join(str(p) for p in table["pages"])
            print(
                f"      sheet '{ws.title}': {len(matrix) - 1} rows x {len(matrix[0])} cols"
                f" (page {pages})"
            )

    if not wb.sheetnames:
        print("Nothing extracted - no workbook written.")
        return 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    print(
        f"\nWrote {out_path}  ({out_path.stat().st_size / 1024:.0f} KB)"
        f"\n  sheets: {len(wb.sheetnames)}   data rows: {total_rows}"
        f"   footer rows dropped: {total_footers}"
    )
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Extract PDF tables into a clean Excel workbook.")
    parser.add_argument("inputs", nargs="*", help="PDF files or globs (default: sample_data/*.pdf)")
    parser.add_argument("-o", "--out", default="output/tables.xlsx", help="output .xlsx path")
    args = parser.parse_args(argv)

    if args.inputs:
        paths: list[Path] = []
        for pattern in args.inputs:
            matches = sorted(Path().glob(pattern)) if any(c in pattern for c in "*?[") else []
            paths.extend(matches or [Path(pattern)])
    else:
        paths = sorted((Path(__file__).parent / "sample_data").glob("*.pdf"))

    paths = [p for p in paths if p.suffix.lower() == ".pdf"]
    if not paths:
        print("No PDF found. Run make_sample_pdf.py first.")
        return 2
    missing = [p for p in paths if not p.exists()]
    if missing:
        print("Missing: " + ", ".join(str(p) for p in missing))
        return 2

    print(f"Reading {len(paths)} PDF file(s)")
    return run(paths, Path(args.out))


if __name__ == "__main__":
    sys.exit(main())

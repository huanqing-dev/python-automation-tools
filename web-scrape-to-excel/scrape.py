#!/usr/bin/env python3
"""Config-driven web scraper that writes a clean Excel workbook.

The point is not "one script for one site" - clients usually want a listing page
turned into a spreadsheet, and they want it to keep working next month. So the
site-specific parts live in a JSON recipe (selectors, pagination, field types)
and this file stays generic.

Handles the things that actually break scrapers in the wild:
  * wrong text encoding (that "Â£51.77" mojibake) - decoded from the response itself
  * pagination via a "next" link, with a hard page cap
  * polite rate limiting and retries with backoff
  * duplicate rows across pages
  * prices like "£51.77", "1,234.50", "12,90 €" -> real numbers
  * one row per item, every row tagged with the page it came from

Usage
    python scrape.py                          # uses sites/books.json, 3 pages
    python scrape.py sites/books.json --pages 5 -o output/books.xlsx
    python scrape.py sites/books.json --details   # also fetch each detail page

Requires: requests, beautifulsoup4, lxml, openpyxl
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "en-GB,en;q=0.9",
}
NUM_RE = re.compile(r"-?\d[\d\s.,]*")


# --------------------------------------------------------------------------- #
# recipe
# --------------------------------------------------------------------------- #
@dataclass
class Recipe:
    name: str
    start_url: str
    item_selector: str
    fields: dict[str, str]
    next_selector: str | None = None
    detail: dict[str, str] = field(default_factory=dict)
    numeric: list[str] = field(default_factory=list)
    dedupe_on: list[str] = field(default_factory=list)

    @classmethod
    def load(cls, path: Path) -> "Recipe":
        data = json.loads(path.read_text())
        return cls(
            name=data["name"],
            start_url=data["start_url"],
            item_selector=data["item_selector"],
            fields=data["fields"],
            next_selector=data.get("next_selector"),
            detail=data.get("detail", {}),
            numeric=data.get("numeric", []),
            dedupe_on=data.get("dedupe_on", []),
        )


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def clean(text: str) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", text or "")).strip()


def pick(node, spec: str) -> str:
    """`css` -> text, `css@attr` -> attribute value, empty spec -> the node itself."""
    selector, _, attr = spec.partition("@")
    element = node.select_one(selector) if selector else node
    if element is None:
        return ""
    if attr:
        value = element.get(attr, "")
        if isinstance(value, list):  # class="star-rating Three" -> "Three"
            used = {c for part in selector.split(".")[1:] for c in part.split()}
            value = " ".join(v for v in value if v not in used)
        return clean(str(value))
    return clean(element.get_text(" "))


def to_number(text: str):
    """'£51.77' -> 51.77, '1,234.50' -> 1234.5, '12,90 €' -> 12.9, else None."""
    if not text:
        return None
    match = NUM_RE.search(text)
    if not match:
        return None
    raw = match.group(0).strip()
    # comma as decimal separator (European format) when there is no dot
    if "," in raw and "." not in raw:
        raw = raw.replace(".", "").replace(",", ".")
    else:
        raw = raw.replace(",", "")
    try:
        return float(raw)
    except ValueError:
        return None


def fetch(session, url, retries=3, delay=0.5, timeout=20):
    last = None
    for attempt in range(retries):
        try:
            response = session.get(url, timeout=timeout)
            if response.status_code == 200:
                # Without this, a page served as text/html with no charset header
                # is decoded as latin-1 and every £ turns into "Â£".
                response.encoding = response.apparent_encoding or "utf-8"
                return response.text
            last = f"HTTP {response.status_code}"
        except requests.RequestException as exc:
            last = str(exc)
        time.sleep(delay * (2 ** attempt))
    raise RuntimeError(f"{url} failed after {retries} tries: {last}")


# --------------------------------------------------------------------------- #
# crawl
# --------------------------------------------------------------------------- #
def crawl(recipe: Recipe, max_pages: int, delay: float, with_details: bool, log=print):
    session = requests.Session()
    session.headers.update(HEADERS)

    rows: list[dict] = []
    seen: set[tuple] = set()
    url, page, duplicates, failed = recipe.start_url, 0, 0, 0
    pages_done = 0

    while url and pages_done < max_pages:
        page += 1
        try:
            html = fetch(session, url, delay=delay)
        except RuntimeError as exc:
            log(f"  ! page {page}: {exc}")
            break
        pages_done += 1
        soup = BeautifulSoup(html, "lxml")
        items = soup.select(recipe.item_selector)
        log(f"  page {page}: {len(items)} items  {url}")

        for node in items:
            row = {name: pick(node, spec) for name, spec in recipe.fields.items()}
            if row.get("url"):
                row["url"] = urljoin(url, row["url"])

            if with_details and recipe.detail:
                detail_url = row.get("url") or ""
                if detail_url:
                    try:
                        detail_html = fetch(session, urljoin(url, detail_url), delay=delay)
                        detail_soup = BeautifulSoup(detail_html, "lxml")
                        for name, spec in recipe.detail.items():
                            row[name] = pick(detail_soup, spec)
                    except RuntimeError as exc:
                        failed += 1
                        log(f"    ! detail: {exc}")

            key = tuple(row.get(k, "") for k in recipe.dedupe_on) if recipe.dedupe_on else None
            if key:
                if key in seen:
                    duplicates += 1
                    continue
                seen.add(key)

            row["_page"] = page
            rows.append(row)
            time.sleep(delay / 2)

        if not recipe.next_selector:
            break
        next_link = soup.select_one(recipe.next_selector)
        url = urljoin(url, next_link["href"]) if next_link and next_link.get("href") else None

    return rows, {"pages": pages_done, "duplicates": duplicates, "failed_details": failed}


# --------------------------------------------------------------------------- #
# workbook
# --------------------------------------------------------------------------- #
HEADER_FILL = PatternFill("solid", fgColor="EEF2F7")


def write_workbook(rows: list[dict], recipe: Recipe, out_path: Path, with_details: bool) -> list[str]:
    columns = list(recipe.fields)
    if with_details:
        columns += [c for c in recipe.detail if c not in columns]
    columns = [c for c in columns if c != "_page"] + ["_page"]

    wb = Workbook()
    ws = wb.active
    ws.title = re.sub(r"[\[\]\*\?/\\:]", "-", recipe.name)[:31]
    ws.append(columns)

    numeric = set(recipe.numeric)
    for row in rows:
        ws.append([row.get(c, "") for c in columns])

    for index, name in enumerate(columns, start=1):
        letter = get_column_letter(index)
        if name in numeric:
            for r in range(2, len(rows) + 2):
                cell = ws.cell(row=r, column=index)
                value = to_number(str(cell.value))
                cell.value = value if value is not None else cell.value
                if value is not None:
                    cell.number_format = "#,##0.00"
        width = max((len(str(ws.cell(row=r, column=index).value or "")) for r in range(1, len(rows) + 2)),
                    default=8)
        ws.column_dimensions[letter].width = min(max(9, width + 2), 60)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return columns


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Scrape a listing site into Excel using a JSON recipe.")
    parser.add_argument("recipe", nargs="?", default="sites/books.json")
    parser.add_argument("-o", "--out", default="output/scraped.xlsx")
    parser.add_argument("--pages", type=int, default=3, help="hard cap on pages (default 3, be polite)")
    parser.add_argument("--delay", type=float, default=0.5, help="seconds between requests")
    parser.add_argument("--details", action="store_true", help="also fetch each detail page")
    args = parser.parse_args(argv)

    recipe_path = Path(args.recipe)
    if not recipe_path.is_absolute() and not recipe_path.exists():
        recipe_path = Path(__file__).parent / args.recipe
    if not recipe_path.exists():
        print(f"recipe not found: {recipe_path}")
        return 2

    recipe = Recipe.load(recipe_path)
    print(f"Scraping {recipe.name} - max {args.pages} page(s), {args.delay}s delay")
    rows, stats = crawl(recipe, args.pages, args.delay, args.details)

    if not rows:
        print("No rows scraped - nothing written.")
        return 1

    out_path = Path(args.out)
    columns = write_workbook(rows, recipe, out_path, args.details)
    print(
        f"\nWrote {out_path}  ({out_path.stat().st_size / 1024:.0f} KB)"
        f"\n  rows: {len(rows)}   pages: {stats['pages']}   duplicates skipped: {stats['duplicates']}"
        f"   detail failures: {stats['failed_details']}"
        f"\n  columns: {', '.join(columns)}"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())

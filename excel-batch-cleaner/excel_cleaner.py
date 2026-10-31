"""
Excel Batch Cleaner v2

Merge multiple Excel workbooks (any number of sheets) into one clean
master file.

Usage:
    python excel_cleaner.py
    python excel_cleaner.py --input ./my_folder --output ./result/clean.xlsx
"""

import argparse
from pathlib import Path

import pandas as pd

EXCEL_PATTERNS = ("*.xlsx", "*.xls", "*.XLSX", "*.XLS")
TAG_COLUMNS = ("source_file", "source_sheet")
DEFAULT_OUTPUT_NAME = "cleaned_master.xlsx"


def find_workbooks(folder: Path, skip: Path) -> list[Path]:
    """Find Excel files (case-insensitive), skipping lock files and our own output."""
    found: dict[str, Path] = {}
    for pattern in EXCEL_PATTERNS:
        for f in sorted(folder.glob(pattern)):
            if f.name.startswith("~$"):  # Excel temp/lock file
                continue
            if f.name.lower() == DEFAULT_OUTPUT_NAME:
                continue  # never re-ingest our own default output
            if f.resolve() == skip.resolve():
                continue  # don't re-ingest a custom output on re-runs
            found.setdefault(f.name.lower(), f)
    return sorted(found.values(), key=lambda p: p.name.lower())


def _text_columns(df: pd.DataFrame) -> list[str]:
    """Text columns on pandas 2.x (object dtype) and pandas 3+/4 (str dtype)."""
    try:
        return list(df.select_dtypes(include="str").columns)
    except TypeError:
        # pandas 2.x rejects the "str" selector entirely
        return list(df.select_dtypes(include="object").columns)


def strip_cells(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace from headers and text cells; numbers are untouched."""
    df.columns = [str(c).strip() for c in df.columns]
    for col in _text_columns(df):
        df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def read_workbook(path: Path) -> tuple[pd.DataFrame, int]:
    """Read every sheet of a workbook, clean and stack them.

    Returns (cleaned frame, raw row count across all sheets).
    """
    sheets = pd.read_excel(path, sheet_name=None)
    raw_rows = sum(len(df) for df in sheets.values())
    frames = []
    for sheet_name, df in sheets.items():
        df = strip_cells(df)
        df = df.dropna(how="all").drop_duplicates().reset_index(drop=True)
        if df.empty:
            continue
        df["source_sheet"] = str(sheet_name)
        frames.append(df)
    if not frames:
        return pd.DataFrame(), raw_rows
    return pd.concat(frames, ignore_index=True), raw_rows


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Merge messy Excel workbooks into one clean master file.")
    p.add_argument("--input", default="sample_data",
                   help="folder with Excel files (default: sample_data)")
    p.add_argument("--output", default=None,
                   help="output .xlsx path (default: <input>/cleaned_master.xlsx)")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    in_dir = Path(args.input)
    out_path = Path(args.output) if args.output else in_dir / "cleaned_master.xlsx"

    files = find_workbooks(in_dir, skip=out_path)
    if not files:
        print(f"No Excel files found in {in_dir}/")
        return

    loaded: list[tuple[Path, pd.DataFrame, int, set[str]]] = []
    failures: list[tuple[str, str]] = []

    for f in files:
        try:
            df, raw_rows = read_workbook(f)
        except Exception as exc:  # corrupt file: skip it, don't kill the batch
            failures.append((f.name, f"unreadable ({exc})"))
            continue
        if df.empty:
            failures.append((f.name, "no usable data"))
            continue
        data_cols = {c for c in df.columns if c not in TAG_COLUMNS}
        loaded.append((f, df, raw_rows, data_cols))

    # Column layout is decided by the most common shape, so one odd file
    # cannot define (or poison) the schema for the whole batch.
    counts: dict[frozenset[str], int] = {}
    for _, _, _, cols in loaded:
        key = frozenset(cols)
        counts[key] = counts.get(key, 0) + 1
    reference_cols: set[str] | None = None
    if counts:
        reference = max(counts.items(), key=lambda kv: kv[1])[0]
        reference_cols = set(reference)
        if len(counts) > 1:
            print(f"Note: {len(counts)} distinct column layouts found; "
                  f"using the most common one ({counts[reference]} file(s)).")

    frames: list[pd.DataFrame] = []
    total_before = 0
    for f, df, raw_rows, data_cols in loaded:
        if data_cols != reference_cols:
            failures.append((f.name, "column mismatch: "
                             f"expected {sorted(reference_cols or [])}, "
                             f"got {sorted(data_cols)}"))
            continue
        total_before += raw_rows
        df["source_file"] = f.name
        frames.append(df)
        print(f"  {f.name}: {len(df)} clean rows")

    if not frames:
        print("Nothing to merge: every file was skipped.")
    else:
        master = pd.concat(frames, ignore_index=True)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        master.to_excel(out_path, index=False)
        print(f"\nDone! {len(frames)} file(s) -> {out_path}")
        print(f"Rows: {total_before} -> {len(master)} "
              f"({total_before - len(master)} junk/duplicate rows removed)")

    if failures:
        print(f"\nSkipped {len(failures)} file(s):")
        for name, reason in failures:
            print(f"  {name}: {reason}")


if __name__ == "__main__":
    main()

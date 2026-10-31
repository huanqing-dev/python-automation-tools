#!/usr/bin/env python3
"""Turn a spreadsheet into a finished PDF report - charts, KPIs and a table.

The recurring freelance job this replaces: "every month I copy these numbers into
a Word/Canva template and export a PDF". Here the layout is code, so the monthly
run is one command.

What it does:
  * reads any sheet, picks up numbers/dates by type instead of by column letter
  * computes KPIs (total, average, best month, change vs previous period)
  * draws two charts with matplotlib and lays everything out with reportlab
  * writes a multi-page A4 PDF with a header, KPI band, table and footer
  * --brand lets a client drop in their own name/colour without touching code

Usage
    python report.py                            # demo data -> output/monthly-report.pdf
    python report.py data/sales.xlsx --brand "Acme Studio" --color "#1a73e8"
    python report.py data/sales.xlsx -o out/q3.pdf --title "Q3 2026 revenue"

Requires: pandas, openpyxl, matplotlib, reportlab
"""

from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # no display needed, and no GUI window pops up
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # noqa: E402
from reportlab.lib.units import mm  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

PAGE_W, PAGE_H = A4


# --------------------------------------------------------------------------- #
# data
# --------------------------------------------------------------------------- #
def load(path: Path | None) -> pd.DataFrame:
    if path is None:
        return demo_frame()
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path)
    return pd.read_csv(path)


def demo_frame() -> pd.DataFrame:
    """Twelve months of plausible numbers so the demo runs with zero setup."""
    months = ["2026-%02d" % m for m in range(1, 13)]
    revenue = [18400, 19250, 22100, 20800, 24600, 26100,
               24900, 28400, 31200, 29800, 33500, 36900]
    orders = [142, 151, 168, 159, 181, 194, 186, 205, 221, 212, 234, 251]
    return pd.DataFrame({"month": months, "revenue": revenue, "orders": orders})


def pick_columns(df: pd.DataFrame):
    """Find a label column and the numeric columns without hard-coding names."""
    numeric = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    label = next((c for c in df.columns if c not in numeric), df.columns[0])
    if not numeric:
        raise SystemExit("No numeric column found - nothing to report on.")
    return label, numeric


def kpis(df: pd.DataFrame, label: str, value: str) -> dict:
    series = df[value].dropna()
    best_row = df.loc[series.idxmax()]
    change = None
    if len(series) >= 2 and series.iloc[-2]:
        change = (series.iloc[-1] - series.iloc[-2]) / abs(series.iloc[-2]) * 100
    return {
        "total": float(series.sum()),
        "average": float(series.mean()),
        "best_label": str(best_row[label]),
        "best_value": float(series.max()),
        "latest": float(series.iloc[-1]) if len(series) else 0.0,
        "change": change,
        "count": len(series),
    }


# --------------------------------------------------------------------------- #
# charts
# --------------------------------------------------------------------------- #
def bar_chart(df, label, value, path: Path, accent: str) -> Path:
    fig, ax = plt.subplots(figsize=(7.4, 2.6), dpi=200)
    ax.bar([str(v) for v in df[label]], df[value], color=accent, width=0.65)
    ax.set_ylabel(value.replace("_", " ").title(), fontsize=8)
    ax.tick_params(axis="x", labelrotation=45, labelsize=7)
    ax.tick_params(axis="y", labelsize=7)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.grid(axis="y", alpha=0.25, linewidth=0.6)
    ax.set_axisbelow(True)
    fig.tight_layout()
    fig.savefig(path, transparent=False, facecolor="white")
    plt.close(fig)
    return path


def line_chart(df, label, value, path: Path, accent: str) -> Path:
    fig, ax = plt.subplots(figsize=(7.4, 2.2), dpi=200)
    ax.plot([str(v) for v in df[label]], df[value], color=accent, linewidth=2, marker="o", markersize=3)
    ax.fill_between(range(len(df)), df[value], color=accent, alpha=0.12)
    ax.tick_params(axis="x", labelrotation=45, labelsize=7)
    ax.tick_params(axis="y", labelsize=7)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.grid(axis="y", alpha=0.25, linewidth=0.6)
    ax.set_axisbelow(True)
    fig.tight_layout()
    fig.savefig(path, facecolor="white")
    plt.close(fig)
    return path


# --------------------------------------------------------------------------- #
# pdf
# --------------------------------------------------------------------------- #
def build_pdf(df: pd.DataFrame, label: str, values: list[str], out: Path,
              title: str, brand: str, accent: str) -> dict:
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=styles["Title"], fontSize=22, spaceAfter=2,
                        textColor=colors.HexColor("#20262e"), alignment=0)
    sub = ParagraphStyle("sub", parent=styles["Normal"], fontSize=9,
                         textColor=colors.HexColor("#6c7682"), spaceAfter=10)
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], fontSize=12.5,
                        textColor=colors.HexColor("#20262e"), spaceBefore=10, spaceAfter=4)

    doc = SimpleDocTemplate(str(out), pagesize=A4,
                            leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=16 * mm, bottomMargin=16 * mm,
                            title=title, author=brand)
    story = [Paragraph(title, h1),
             Paragraph(f"{brand} &nbsp;|&nbsp; generated {date.today():%d %b %Y} "
                       f"&nbsp;|&nbsp; {len(df)} rows", sub)]

    value = values[0]
    stats = kpis(df, label, value)
    kpi_cells = [
        ("Total", f"{stats['total']:,.0f}"),
        ("Average", f"{stats['average']:,.1f}"),
        (f"Best ({stats['best_label']})", f"{stats['best_value']:,.0f}"),
        ("Latest", f"{stats['latest']:,.0f}"),
    ]
    if stats["change"] is not None:
        arrow = "▲" if stats["change"] >= 0 else "▼"
        kpi_cells.append(("vs previous", f"{arrow} {abs(stats['change']):.1f}%"))

    table = Table([[c[1] for c in kpi_cells], [c[0] for c in kpi_cells]],
                  colWidths=[(PAGE_W - 36 * mm) / len(kpi_cells)] * len(kpi_cells))
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f2f5f9")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#dde3ea")),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#e6ebf1")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 14),
        ("FONTSIZE", (0, 1), (-1, 1), 7.5),
        ("TEXTCOLOR", (0, 1), (-1, 1), colors.HexColor("#6c7682")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(accent)),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story += [table, Spacer(1, 8)]

    tmp = out.parent
    tmp.mkdir(parents=True, exist_ok=True)
    chart_w = PAGE_W - 36 * mm
    story += [Paragraph("Trend", h2)]
    bar = bar_chart(df, label, value, tmp / "_bar.png", accent)
    story += [Image(str(bar), width=chart_w, height=chart_w * 2.6 / 7.4)]
    story += [Paragraph("Cumulative", h2)]
    cum = df.copy()
    cum[value] = cum[value].cumsum()
    line = line_chart(cum, label, value, tmp / "_line.png", accent)
    story += [Image(str(line), width=chart_w, height=chart_w * 2.2 / 7.4)]

    # data table on its own page
    story += [PageBreak(), Paragraph("Data", h2)]
    head = [str(c).replace("_", " ").title() for c in df.columns]
    body = []
    for _, row in df.iterrows():
        cells = []
        for col, v in zip(df.columns, row):
            if isinstance(v, (int, float)):
                cells.append(f"{v:,.0f}" if float(v).is_integer() else f"{v:,.2f}")
            else:
                cells.append(str(v))
        body.append(cells)
    data_table = Table([head] + body, repeatRows=1,
                       colWidths=[(PAGE_W - 36 * mm) / len(head)] * len(head))
    data_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eef2f7")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#dde3ea")),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#fafbfc")]),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(KeepTogether(data_table))

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.HexColor("#9aa5b1"))
        canvas.drawString(18 * mm, 10 * mm, f"{brand} - {title}")
        canvas.drawRightString(PAGE_W - 18 * mm, 10 * mm, f"Page {doc_.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    for leftover in ("_bar.png", "_line.png"):
        (tmp / leftover).unlink(missing_ok=True)
    return stats


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Build a PDF report from a spreadsheet.")
    parser.add_argument("data", nargs="?", help="xlsx/csv file (omit for demo data)")
    parser.add_argument("-o", "--out", default="output/monthly-report.pdf")
    parser.add_argument("--title", default="Monthly performance report")
    parser.add_argument("--brand", default="Acme Studio")
    parser.add_argument("--color", default="#1a73e8", help="accent colour, e.g. #1a73e8")
    args = parser.parse_args(argv)

    df = load(Path(args.data) if args.data else None)
    label, numeric = pick_columns(df)
    out = Path(args.out)
    stats = build_pdf(df, label, numeric, out, args.title, args.brand, args.color)

    print(f"Wrote {out}  ({out.stat().st_size / 1024:.0f} KB)")
    print(f"  rows: {len(df)}   label column: {label}   numeric: {', '.join(numeric)}")
    print(f"  total {stats['total']:,.0f}   average {stats['average']:,.1f}   "
          f"best {stats['best_label']} ({stats['best_value']:,.0f})")
    return 0


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
"""Build a realistic sample PDF so the extractor can be tested without your own files.

The generated file mimics the annoying parts of real-world PDFs:
  * a ruled table that spans 3 pages with the header repeated on every page
  * numbers with thousands separators, currency symbols and (parenthesised) negatives
  * a percentage column
  * a per-page footer ("Page 1 of 3") that must NOT end up in the spreadsheet
  * a second, border-less table on the last page

Run:  python make_sample_pdf.py
"""

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

OUT = Path(__file__).parent / "sample_data" / "q3-expenses.pdf"

HEADER = ["Date", "Description", "Qty", "Unit price", "VAT %", "Amount"]

ROWS = [
    ["2026-07-02", "Server hosting - July", "1", "¥ 12,400.00", "6%", "¥ 12,400.00"],
    ["2026-07-05", "Design software (annual)", "3", "$ 249.90", "0%", "$ 749.70"],
    ["2026-07-11", "Contractor - data entry", "38", "¥ 65.00", "6%", "¥ 2,470.00"],
    ["2026-07-18", "Ad spend - search", "1", "$ 1,820.35", "0%", "$ 1,820.35"],
    ["2026-07-23", "Refund - duplicate invoice", "1", "(¥ 480.00)", "6%", "(¥ 480.00)"],
    ["2026-08-01", "Server hosting - August", "1", "¥ 12,400.00", "6%", "¥ 12,400.00"],
    ["2026-08-04", "Translation - DE/EN", "12", "€ 45.00", "0%", "€ 540.00"],
    ["2026-08-09", "Office chairs", "4", "¥ 1,299.00", "13%", "¥ 5,196.00"],
    ["2026-08-14", "Bank fees", "1", "¥ 38.50", "0%", "¥ 38.50"],
    ["2026-08-20", "Ad spend - social", "1", "$ 960.00", "0%", "$ 960.00"],
    ["2026-08-27", "Domain renewals", "6", "$ 12.99", "0%", "$ 77.94"],
    ["2026-09-03", "Server hosting - September", "1", "¥ 12,400.00", "6%", "¥ 12,400.00"],
    ["2026-09-08", "Contractor - QA pass", "22", "¥ 80.00", "6%", "¥ 1,760.00"],
    ["2026-09-15", "Accounting software", "1", "¥ 3,588.00", "13%", "¥ 3,588.00"],
    ["2026-09-21", "Travel - client meeting", "1", "¥ 1,260.00", "6%", "¥ 1,260.00"],
    ["2026-09-28", "Monitoring service", "1", "$ 29.00", "0%", "$ 29.00"],
    ["2026-07-03", "Email service - July", "1", "$ 42.00", "0%", "$ 42.00"],
    ["2026-07-06", "Stock photography", "10", "$ 9.00", "0%", "$ 90.00"],
    ["2026-07-08", "Courier - contracts", "2", "¥ 68.00", "6%", "¥ 136.00"],
    ["2026-07-13", "Cloud storage top-up", "1", "$ 120.00", "0%", "$ 120.00"],
    ["2026-07-15", "Freelance illustration", "1", "¥ 3,200.00", "6%", "¥ 3,200.00"],
    ["2026-07-20", "Print - business cards", "250", "¥ 1.20", "13%", "¥ 300.00"],
    ["2026-07-27", "API credits", "1", "$ 300.00", "0%", "$ 300.00"],
    ["2026-08-02", "Email service - August", "1", "$ 42.00", "0%", "$ 42.00"],
    ["2026-08-06", "Team lunch - kickoff", "6", "¥ 88.00", "6%", "¥ 528.00"],
    ["2026-08-11", "Stock photography", "14", "$ 9.00", "0%", "$ 126.00"],
    ["2026-08-13", "Laptop repair", "1", "¥ 1,480.00", "13%", "¥ 1,480.00"],
    ["2026-08-17", "Ad spend - retargeting", "1", "$ 640.00", "0%", "$ 640.00"],
    ["2026-08-22", "Courier - samples", "3", "¥ 68.00", "6%", "¥ 204.00"],
    ["2026-08-29", "API credits", "1", "$ 260.00", "0%", "$ 260.00"],
    ["2026-09-01", "Email service - September", "1", "$ 42.00", "0%", "$ 42.00"],
    ["2026-09-05", "Refund - cancelled licence", "1", "($ 249.90)", "0%", "($ 249.90)"],
    ["2026-09-09", "Courier - signed copies", "4", "¥ 68.00", "6%", "¥ 272.00"],
    ["2026-09-12", "Contractor - landing page", "1", "¥ 2,800.00", "6%", "¥ 2,800.00"],
    ["2026-09-17", "Stock photography", "8", "$ 9.00", "0%", "$ 72.00"],
    ["2026-09-19", "Bank fees", "1", "¥ 38.50", "0%", "¥ 38.50"],
    ["2026-09-24", "Ad spend - search", "1", "$ 1,140.00", "0%", "$ 1,140.00"],
    ["2026-09-26", "Backup drive", "2", "¥ 899.00", "13%", "¥ 1,798.00"],
    ["2026-09-30", "API credits", "1", "$ 410.00", "0%", "$ 410.00"],
]


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)
    canvas.drawString(20 * mm, 12 * mm, "Acme Studio Ltd - internal report - do not distribute")
    canvas.drawRightString(190 * mm, 12 * mm, f"Page {doc.page} of 3")
    canvas.restoreState()


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(OUT), pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=18 * mm, bottomMargin=20 * mm,
        title="Q3 expenses", author="Acme Studio Ltd",
    )
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Q3 2026 - Expense detail", styles["Title"]),
        Paragraph(
            "All amounts are in the original billing currency. Negative amounts are refunds.",
            styles["Normal"],
        ),
        Spacer(1, 8 * mm),
    ]

    data = [HEADER] + ROWS
    table = Table(data, repeatRows=1, colWidths=[24 * mm, 52 * mm, 13 * mm, 26 * mm, 16 * mm, 30 * mm])
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#9aa5b1")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eef2f7")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(table)

    story += [
        Spacer(1, 10 * mm),
        Paragraph("Summary by category", styles["Heading2"]),
        Spacer(1, 3 * mm),
    ]

    # A border-less table, centred columns - a very common layout that line-based
    # extraction cannot see, so the extractor needs its text-based fallback.
    summary = [
        ["Category", "Items", "Total"],
        ["Infrastructure", "3", "37,200.00"],
        ["Software", "3", "4,587.80"],
        ["Contractors", "2", "4,230.00"],
        ["Marketing", "2", "2,780.35"],
    ]
    t2 = Table(summary, colWidths=[50 * mm, 25 * mm, 35 * mm], hAlign="CENTER")
    t2.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.black),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(t2)

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    print(f"wrote {OUT} ({OUT.stat().st_size/1024:.0f} KB)")


if __name__ == "__main__":
    build()

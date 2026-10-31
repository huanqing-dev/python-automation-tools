# excel-to-pdf-report

Turn a spreadsheet into a finished PDF report — KPI band, two charts, full data table, footer.

The job this replaces: every month someone copies numbers into a Word/Canva
template, fiddles with the layout, exports a PDF and sends it. The layout here is
code, so next month is **one command**.

## What you get

| Page | Content |
|---|---|
| 1 | Title + brand line, KPI band (total / average / best period / latest / change vs previous), a bar chart of the trend, a cumulative line chart |
| 2 | The full data table, zebra-striped, repeating header if it spills onto more pages |
| every page | Footer with brand, report title and page number |

Numbers and dates are picked up **by type**, not by hard-coded column names —
point it at a different file and it still works.

## Install

```bash
pip install -r requirements.txt
```

## Use

```bash
python report.py                                     # built-in demo data
python report.py data/sales.xlsx --brand "Acme Studio" --color "#1a73e8"
python report.py data/sales.xlsx -o out/q3.pdf --title "Q3 2026 revenue"
```

Demo run:

```
Wrote output/monthly-report.pdf  (85 KB)
  rows: 12   label column: month   numeric: revenue, orders
  total 315,950   average 26,329.2   best 2026-12 (36,900)
```

## Client branding without code

`--brand` changes the name in the header and footer, `--color` changes the
accent colour of the KPI band and charts. A client can have "their" report in
one flag.

## Limits (stated up front)

- One label column + numeric columns; wide "one row per date" pivots should be
  melted first (I can do that as part of the job).
- Charts are matplotlib defaults styled minimally — pixel-perfect brand design
  is a separate task.
- No logins, no email sending, no scheduling in this script; wiring it to cron
  or a shared folder watcher is a small add-on.

## License

MIT — see [../LICENSE](../LICENSE).

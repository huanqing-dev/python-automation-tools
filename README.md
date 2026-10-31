# Python automation tools

Small, tested tools for repetitive spreadsheet and document work.
Each folder is a standalone project with its own README, sample data and tests.

## Projects

### 1. [excel-batch-cleaner](excel-batch-cleaner/)
Merge a folder of messy Excel workbooks into one clean master file.

- Reads every sheet, not just the first
- Strips stray spaces, drops blank rows and exact duplicates
- Validates column layouts — mismatched or corrupt files are reported, never silently dropped
- Tags each row with its source file and sheet
- Prints a before/after report

![before/after](docs/excel-before-after.png)

```bash
cd excel-batch-cleaner
pip install -r requirements.txt
python make_sample_data.py
python excel_cleaner.py
```

## About

I build small automation tools in Python (pandas, openpyxl, pdfplumber).
Fixed scope, fixed price, tested against real edge cases before delivery.

Available for freelance work — data cleaning, Excel/CSV automation, report generation.

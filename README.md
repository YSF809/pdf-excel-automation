# PDF & Excel Automation

A Python desktop app that extracts invoice fields from text-based PDF files, cleans CSV data, and writes a formatted Excel report.

## Features

- Select PDF and CSV files through a Tkinter interface.
- Extract invoice number, date, customer, subtotal, VAT, and total from supported invoice text layouts.
- Mark incomplete PDF records as `Needs Review`.
- Trim CSV headers and text, remove empty rows and duplicate rows.
- Create Excel sheets for PDF invoices, CSV data, and processing errors when applicable.
- Format headers, filters, frozen rows, and SAR amounts. Raw PDF text is stored in a hidden column.

## Setup on Windows

Install Python with Tkinter support, then open a terminal in this project folder:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe GUI.py
```

Tkinter is included with standard Windows Python installations; it is not a pip requirement.

## Use the desktop app

1. Click **Add PDF / CSV Files**.
2. Click **Choose Output Folder**.
3. Click **Generate Excel Report**.
4. Open the timestamped `report_YYYYMMDD_HHMMSS.xlsx` file in the selected folder.

## Use the processing module from the terminal

Place local PDF and CSV files in `input_files`, then run from this project folder:

```powershell
.\.venv\Scripts\python.exe -c "import main; main.generate_excel()"
```

Reports are written to `output`. Running `main.py` directly does not generate a report; use the command above or the desktop app.

## Limitations and data privacy

PDF parsing uses text extraction and patterns for supported English invoice layouts. Scanned/image-only PDFs require OCR, which this project does not provide. Other invoice layouts may need pattern changes. Review extracted amounts and any records marked `Needs Review` before using the report.

Reports can contain customer information, source filenames, and raw invoice text, including in hidden columns. Keep input files and reports private. The `.gitignore` excludes local input/output folders, PDF/CSV files, Excel workbooks, environments, and common credential files. Only publish reviewed source code and documentation.

## Project files

- `GUI.py`: desktop interface and background processing.
- `main.py`: invoice extraction, CSV cleaning, report generation, and formatting.
- `requirements.txt`: third-party Python dependencies.
- `.gitignore`: excludes private data and local/generated files.

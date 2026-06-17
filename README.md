# QA Audit CSV to Excel & PDF Dashboard Converter

Python automation tool that converts raw QA audit logs (CSV) into a formatted Excel dashboard and PDF report — no Microsoft Excel or LibreOffice required.

## Project Structure

```
.
├── run.py              # Entry point: CSV → Excel → PDF in one command
├── format_excel.py     # Core logic: CSV → formatted Excel + PNG chart + PDF
├── excel_to_pdf.py     # Standalone Excel → PDF converter
├── input/              # Place your CSV file here
└── output/             # Generated files land here
```

## Features

- **One command workflow:** `python run.py` handles everything end-to-end.
- **CSV → Excel:** Converts QA audit CSV logs into a fully formatted `.xlsx` workbook.
- **Conditional Formatting:** Color-codes rows based on AI Decision values.
- **Auto-Sizing:** Adjusts column widths and wraps long text fields.
- **Dashboard Sheet:** Summary KPI cards + pie chart in a separate sheet.
- **Chart Export:** Saves the decisions chart as a `.png` image.
- **PDF Export:** Converts the Excel output to PDF automatically (no Excel needed).

## Prerequisites

- Python 3.10+

## Installation

```bash
git clone https://github.com/Al-Oqab/CVS-to-Excel-Tool.git
cd CVS-to-Excel-Tool
python -m pip install -r requirements.txt
```

## Usage

### Full pipeline (recommended)

Place one or more CSV files in the `input/` folder, then run:

```bash
python run.py
```

Each CSV gets its own set of output files in `output/`:
- `filename.xlsx` — formatted Excel report
- `filename_chart.png` — decisions pie chart
- `filename.pdf` — PDF version of the Excel report

### Convert Excel to PDF only

Convert all `.xlsx` files in `output/` at once:

```bash
python excel_to_pdf.py
```

Convert a single file:

```bash
python excel_to_pdf.py output/report.xlsx
```

Custom output path:

```bash
python excel_to_pdf.py output/report.xlsx -o reports/report.pdf
```

Convert a specific folder:

```bash
python excel_to_pdf.py -f some_folder/
```

## Dependencies

| Package    | Purpose                       |
|------------|-------------------------------|
| pandas     | CSV reading & data processing |
| openpyxl   | Excel file creation/reading   |
| matplotlib | Charts & PDF rendering        |
| fpdf2      | PDF report generation         |

Install all at once:

```bash
python -m pip install pandas openpyxl matplotlib fpdf2
```

## License

MIT License

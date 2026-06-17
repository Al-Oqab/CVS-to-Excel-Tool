# QA Audit CSV to Excel & PDF Dashboard Converter

Python automation tool that converts raw QA audit logs (CSV) into a formatted Excel dashboard, and optionally exports them to PDF — no Microsoft Excel or LibreOffice required.

## Project Structure

```
.
├── format_excel.py     # CSV → formatted Excel + PNG chart + PDF report
├── excel_to_pdf.py     # Excel → PDF (standalone converter)
├── output/             # Generated files land here
└── input/              # Place your CSV files here
```

## Features

- **CSV → Excel:** Converts QA audit CSV logs into a fully formatted `.xlsx` workbook.
- **Conditional Formatting:** Color-codes rows based on AI Decision values.
- **Auto-Sizing:** Adjusts column widths and wraps long text fields.
- **Dashboard Sheet:** Summary KPI cards + pie chart in a separate sheet.
- **Chart Export:** Saves the decisions chart as a `.png` image.
- **PDF Export:** Converts any `.xlsx` file to PDF using Python only (no Excel needed).

## Prerequisites

- Python 3.10+

## Installation

```bash
git clone https://github.com/Al-Oqab/CVS-to-Excel-Tool.git
cd CVS-to-Excel-Tool
python -m pip install -r requirements.txt
```

## Usage

### 1. Generate Excel report from CSV

```bash
python format_excel.py
```

Output files are saved to `output/`.

### 2. Convert Excel to PDF

Convert all `.xlsx` files in `output/` at once:

```bash
python excel_to_pdf.py
```

Convert a single file:

```bash
python excel_to_pdf.py output/master_tracker.xlsx
```

Convert a specific folder:

```bash
python excel_to_pdf.py -f some_folder/
```

Custom output path:

```bash
python excel_to_pdf.py output/master_tracker.xlsx -o reports/tracker.pdf
```

## Dependencies

| Package     | Purpose                        |
|-------------|--------------------------------|
| pandas      | CSV reading & data processing  |
| openpyxl    | Excel file creation/reading    |
| matplotlib  | Charts & PDF rendering         |
| fpdf2       | PDF report generation          |

Install all at once:

```bash
python -m pip install pandas openpyxl matplotlib fpdf2
```

## License

MIT License

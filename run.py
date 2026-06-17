"""
Entry point: CSV → Excel → PDF in one command.

    python run.py

Reads all CSV files from the 'input' folder and writes Excel + PNG + PDF
to 'output' for each one.
"""

import os
import sys
from format_excel import generate_audit_report


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(base_dir, "input")
    output_dir = os.path.join(base_dir, "output")

    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    csv_files = sorted(f for f in os.listdir(input_dir) if f.endswith(".csv"))

    if not csv_files:
        print(f"Error: No CSV file found in '{input_dir}'")
        sys.exit(1)

    print(f"Found {len(csv_files)} CSV file(s) to process.\n")

    success, failed = [], []

    for i, csv_file in enumerate(csv_files, 1):
        csv_name     = os.path.splitext(csv_file)[0]
        input_file   = os.path.join(input_dir,  csv_file)
        output_excel = os.path.join(output_dir, f"{csv_name}.xlsx")
        output_image = os.path.join(output_dir, f"{csv_name}_chart.png")
        output_pdf   = os.path.join(output_dir, f"{csv_name}.pdf")

        print(f"[{i}/{len(csv_files)}] Processing: {csv_file}")
        try:
            generate_audit_report(input_file, output_excel, output_image, output_pdf)
            success.append(csv_file)
        except Exception as e:
            print(f"  Error: {e}")
            failed.append(csv_file)
        print()

    print("=" * 50)
    print(f"Done: {len(success)} succeeded, {len(failed)} failed.")
    if failed:
        print("Failed files:")
        for f in failed:
            print(f"  - {f}")


if __name__ == "__main__":
    main()

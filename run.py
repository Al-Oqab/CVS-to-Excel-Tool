"""
Entry point: CSV → Excel → PDF in one command.

    python run.py

Reads the CSV from the 'input' folder and writes Excel + PDF to 'output'.
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

    csv_files = [f for f in os.listdir(input_dir) if f.endswith(".csv")]

    if not csv_files:
        print(f"Error: No CSV file found in '{input_dir}'")
        sys.exit(1)
    if len(csv_files) > 1:
        print(f"Error: Multiple CSV files found: {csv_files}")
        print("Keep only one CSV in the 'input' folder.")
        sys.exit(1)

    csv_name = os.path.splitext(csv_files[0])[0]
    input_file   = os.path.join(input_dir,  csv_files[0])
    output_excel = os.path.join(output_dir, f"{csv_name}.xlsx")
    output_image = os.path.join(output_dir, f"{csv_name}_chart.png")
    output_pdf   = os.path.join(output_dir, f"{csv_name}.pdf")

    print(f"Processing: {csv_files[0]}")
    generate_audit_report(input_file, output_excel, output_image, output_pdf)
    print("\nAll done!")
    print(f"  Excel : {output_excel}")
    print(f"  Chart : {output_image}")
    print(f"  PDF   : {output_pdf}")


if __name__ == "__main__":
    main()

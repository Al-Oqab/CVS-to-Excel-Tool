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
    if len(csv_files) > 1:
        print(f"Error: Multiple CSV files found: {csv_files}")
        print("Keep only one CSV in the 'input' folder.")
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

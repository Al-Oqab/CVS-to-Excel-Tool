import pandas as pd
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import PieChart, Reference

def generate_audit_report(input_csv: str, output_excel: str):
    """
    Reads a CSV file containing QA Audit data and converts it into a formatted Excel
    file with a Dashboard and conditional formatting.
    """
    # 1. Check if the input file exists
    if not os.path.exists(input_csv):
        print(f"Error: The file '{input_csv}' was not found.")
        print("Please ensure the CSV file is in the same directory as the script.")
        sys.exit(1)

    print(f"Reading data from {input_csv}...")
    df = pd.read_csv(input_csv)

    wb = Workbook()
    
    # ------------------ SHEET 1: AUDIT DATA ------------------
    ws_data = wb.active
    ws_data.title = "Audit Data"

    # Insert data
    for r in dataframe_to_rows(df, index=False, header=True):
        ws_data.append(r)

    # Styling Constants
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(left=Side(style='thin', color='E0E0E0'), 
                         right=Side(style='thin', color='E0E0E0'), 
                         top=Side(style='thin', color='E0E0E0'), 
                         bottom=Side(style='thin', color='E0E0E0'))

    # Format Headers
    for cell in ws_data[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Set Column Widths
    column_widths = {'A': 22, 'B': 10, 'C': 12, 'D': 20, 'E': 90}
    for col, width in column_widths.items():
        ws_data.column_dimensions[col].width = width

    # Conditional Formatting Colors mapping for 'AI Decision'
    colors = {
        "NEW_DAMAGE": "E2F0D9",  # Green
        "DISCARD": "FFF2CC",     # Yellow
        "ERROR": "FCE4D6",       # Red
        "MANUAL_REVIEW": "DDEBF7" # Blue
    }
    text_colors = {
        "NEW_DAMAGE": "385723",
        "DISCARD": "7F6000",
        "ERROR": "C00000",
        "MANUAL_REVIEW": "1F4E78"
    }

    # Apply formatting to rows
    # Note: Assuming 'AI Decision' is in the 4th column (Index 3)
    for row in ws_data.iter_rows(min_row=2, max_row=ws_data.max_row):
        status = str(row[3].value).strip() if row[3].value else ""
        fill_color = colors.get(status, "FFFFFF")
        text_color = text_colors.get(status, "000000")
        
        fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        font = Font(color=text_color, bold=True)
        
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=(cell.column_letter == 'E'))
        
        # Color specific cell
        row[3].fill = fill
        row[3].font = font

    # ------------------ SHEET 2: DASHBOARD ------------------
    print("Generating Dashboard...")
    ws_summary = wb.create_sheet(title="Dashboard")
    ws_summary.sheet_view.showGridLines = False

    # Ensure the column exists before counting
    if "AI Decision" in df.columns:
        counts = df["AI Decision"].value_counts()
        ws_summary.append(["AI Decision", "Count"])
        for decision, count in counts.items():
            ws_summary.append([decision, count])

        # Style Summary Table
        for cell in ws_summary[1]:
            cell.fill = header_fill
            cell.font = header_font
            
        for row in ws_summary.iter_rows(min_row=2, max_row=len(counts)+1, min_col=1, max_col=2):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center")

        ws_summary.column_dimensions['A'].width = 20
        ws_summary.column_dimensions['B'].width = 15

        # Add Pie Chart
        pie = PieChart()
        labels = Reference(ws_summary, min_col=1, min_row=2, max_row=len(counts)+1)
        data_ref = Reference(ws_summary, min_col=2, min_row=1, max_row=len(counts)+1)
        pie.add_data(data_ref, titles_from_data=True)
        pie.set_categories(labels)
        pie.title = "AI Decisions Summary"
        ws_summary.add_chart(pie, "D2")
    else:
        print("Warning: 'AI Decision' column not found. Skipping chart generation.")

    # Save File
    wb.save(output_excel)
    print(f"Success! Excel report generated at: {output_excel}")


if __name__ == "__main__":
    # Define your file names here
    INPUT_FILE = "qa_audit_tracker.csv"
    OUTPUT_FILE = "QA_Audit_Presentation.xlsx"
    
    generate_audit_report(INPUT_FILE, OUTPUT_FILE)
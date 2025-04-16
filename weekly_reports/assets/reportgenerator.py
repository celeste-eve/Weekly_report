import pandas as pd
import sys
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import datetime

# date and time for file saving 
D = datetime.datetime.now().strftime("%Y-%m-%d %H_%M")

def generate_summary(input_file, output_dir):
    # Load the spreadsheet
    df = pd.read_excel(input_file, header=2)

    # select the correct sheet based on the largest numeric name
    excel_file = input_file

    sheet_names = pd.ExcelFile(excel_file).sheet_names

    numeric_sheet_names = [name for name in sheet_names if name.isdigit()]

    
    numeric_sheet_numbers = [int(name) for name in numeric_sheet_names]


    if numeric_sheet_numbers:
        max_sheet_number = max(numeric_sheet_numbers)
        max_sheet_name = str(max_sheet_number)

        
        df = pd.read_excel(excel_file, sheet_name=max_sheet_name)

        # Proceed with your data processing
        print(f"Data from sheet '{max_sheet_name}':")
        print(df.head())
    else:
        print("No sheets with purely numeric names found.")


    # Clean headers
    df.columns = df.columns.str.strip()

    # Debug: print the actual column names
    print("Found columns:", df.columns)

    # Now  selecting relevant columns
    relevant_columns = ['Test Types', 'Scheduled', 'InProgress', 'Completed', 'Checked', 'Approved', 'NCF? Restricted numbers', 'Week Approved Increase']

    # Check if all relevant columns exist
    missing_cols = [col for col in relevant_columns if col not in df.columns]
    if missing_cols:
        print(f"Missing columns: {missing_cols}")
        return

    df = df[relevant_columns]

    # define first colum and group by it
    # first col contains project names 
    first_col = df.columns[0]

    grouped = df.groupby(first_col)

    grouped = {name: group.reset_index(drop=True) for name, group in grouped}

   

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python reportgenerator.py <input_file> <output_dir>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_dir = sys.argv[2]
    generate_summary(input_file, output_dir)


# DATA FRAME needs to be replaced with data 
data = {
    ("Test Types"): ['Test Types'],
    ("Scheduled"): [1, 2, 3, 4],
    ("InProgress"): [1, 2, 3, 4],
    ("Completed"): [1, 2, 3, 4],
    ("Checked"): [1, 2, 3, 4],
    ("Approved"): [1, 2, 3, 4],
    ("NCF/Restricted"): [1, 2, 3, 4],
    ("Total"): [1, 2, 3, 4],
    ("Weekly Total completed"): [1, 2, 3, 4],

}
index = ["ASTM Test A", "ISO Test B", "Mechanical", "TOTALS"] # needs to be replaced with tests 

df = pd.DataFrame(data, index=index)

grouped_dfs = df.groupby(level=0).sum()

# Create workbook and sheet
wb = Workbook()
ws = wb.active
ws.title = "Summary Report"

 # Save to output directory
output_path = os.path.join(output_dir, f"weekly_report_summary_{D}.xlsx")
grouped_dfs.to_excel(output_path, index=False)
print(f"Report generated: {output_path}")

# Write the column headers (multi-level)
ws.append(["Test Types"] + [header for col in df.columns for header in col])

# Style header
for col_num, col in enumerate(["Test Types"] + [header for col in df.columns for header in col], 1):
    cell = ws.cell(row=1, column=col_num)
    cell.fill = PatternFill(start_color="E4A8D8", fill_type="solid")
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")


# Write the data rows
for row in dataframe_to_rows(df, index=True, header=False):
    ws.append(row)

# Style entire table (pink background, alignment)
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
    for cell in row:
        cell.fill = PatternFill(start_color="FAD7F6", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Save workbook
wb.save(output_dir)
print(f"Saved to {output_dir}")
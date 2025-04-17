import pandas as pd
import sys
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime

# Get current timestamp
D = datetime.datetime.now().strftime("%Y-%m-%d %H_%M")

def generate_summary(input_file, output_dir):
    # Load sheet with largest numeric name
    xls = pd.ExcelFile(input_file)
    numeric_sheets = [int(name) for name in xls.sheet_names if name.isdigit()]
    
    if not numeric_sheets:
        print("No numeric sheet names found.")
        return
    
    sheet_name = str(max(numeric_sheets))
    df = pd.read_excel(input_file, sheet_name=sheet_name, header=1)


    # Rename columns by position (first becomes 'Project', second becomes 'Test Type')
    df.columns.values[0] = 'Project'
    df.columns.values[1] = 'Test Type'

    print("Column names:")
    print(list(df.columns))

    # print columns found 
    print("Found columns:", df.columns)

    
    # Define columns to keep
    columns_to_keep = ['Project', 'Broad type', 'Test Type', 'Scheduled', 'Checked', 'InProgress', 'Completed', 'Approved', 'Total', 'NCF? Restricted numbers', 'Week Approved Increase']

    # Filter only those columns if they exist
    df = df[[col for col in columns_to_keep if col in df.columns]]

    missing_cols = [col for col in columns_to_keep if col not in df.columns]
    if missing_cols:
        print(f"WARNING: Missing columns in input file: {missing_cols}")


    required_grouping_cols = ['Project', 'Broad type']
    if not all(col in df.columns for col in required_grouping_cols):
        print(f"Missing required columns for grouping: {required_grouping_cols}")
        return

    first_col = 'Project'
    type_col = 'Broad type'
    df['NCF? Restricted numbers'] = pd.to_numeric(df['NCF? Restricted numbers'], errors='coerce').abs()



    # Group by both Project and Type
    grouped = df.groupby([first_col, type_col])

    desired_order = [
    'Project', 'Test Type', 'Broad type',
    'Scheduled', 'InProgress', 'Completed', 'Checked', 'Approved', 'Total',
    'NCF? Restricted numbers', 'Week Approved Increase'
    ]

    # Reorder columns (only those that are actually in df)
    df = df[[col for col in desired_order if col in df.columns]]


    # Output Excel file path
    output_path = os.path.join(output_dir, f"weekly_report_summary_{D}.xlsx")

    # Create Excel writer using openpyxl engine
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for (project, test_type), group in grouped:
            sheet_name = f"{project[:15]}_{test_type[:10]}"  
            print(f"Writing group: {sheet_name}")
            group.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Report generated: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python reportgenerator.py <input_file> <output_dir>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_dir = sys.argv[2]
    generate_summary(input_file, output_dir)

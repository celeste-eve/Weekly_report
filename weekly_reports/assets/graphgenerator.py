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

    #grouped by project 
    grouped = df.groupby([first_col])

    summary_df = grouped.agg({
        'Completed': 'sum',
        'Approved': 'sum',
        'Checked': 'sum',
        'Scheduled': 'sum',
        'InProgress': 'sum',
        'NCF? Restricted numbers': 'sum'
    }).reset_index()

    # Create the final summary columns
    summary_df['Total Completed'] = summary_df['Completed'] + summary_df['Approved'] + summary_df['Checked']
    summary_df['Total Remaining'] = summary_df['Scheduled'] + summary_df['InProgress']

    # Select and rename columns
    final_df = summary_df[[first_col, 'Total Completed', 'Total Remaining', 'NCF? Restricted numbers']]
    final_df = final_df.rename(columns={'NCF? Restricted numbers': 'Total NCF? Restricted numbers'})


    # Set output file path
    output_file = os.path.join(output_dir, f"summary_output_{D}.xlsx")

    # Create a writer object
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write individual dataframes to separate sheets
        summary_df.to_excel(writer, sheet_name='Completed', index=False)
    

    print(f"\n Summary Excel file saved to: {output_file}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python graphgenerator.py <input_file> <output_dir>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_dir = sys.argv[2]
    generate_summary(input_file, output_dir)

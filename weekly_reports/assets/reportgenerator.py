import pandas as pd
import sys
import os

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

        
        print(f"Data from sheet '{max_sheet_name}':")
        print(df.head())
    else:
        print("No sheets with purely numeric names found.")

    # Clean headers
    df.columns = df.columns.str.strip()

    # Debug: print the actual column names
    print("Found columns:", df.columns)

    # Now try selecting relevant columns
    relevant_columns = ['Scheduled', 'InProgress', 'Completed', 'Checked', 'Approved', 'NCF']

    # Check if all relevant columns exist
    missing_cols = [col for col in relevant_columns if col not in df.columns]
    if missing_cols:
        print(f"Missing columns: {missing_cols}")
        return

    df = df[relevant_columns]

    # Group and summarize
    summary = df.groupby(['Project', 'Test Type'], as_index=False).sum(numeric_only=True)

    # Save to output directory
    output_path = os.path.join(output_dir, 'weekly_report_summary.xlsx')
    summary.to_excel(output_path, index=False)
    print(f"Report generated: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python reportgenerator.py <input_file> <output_dir>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_dir = sys.argv[2]
    generate_summary(input_file, output_dir)

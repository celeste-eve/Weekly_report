import pandas as pd
import sys
import os
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter
import datetime
from fpdf import FPDF

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

    # Rename columns
    df.columns.values[0] = 'Project'
    df.columns.values[1] = 'Test Type'

    columns_to_keep = ['Project', 'Broad type', 'Test Type', 'Scheduled', 'Checked', 'InProgress', 'Completed', 'Approved', 'Total', 'NCF? Restricted numbers', 'Week Approved Increase']
    df = df[[col for col in columns_to_keep if col in df.columns]]

    # Cleanup and calculations
    df['NCF? Restricted numbers'] = pd.to_numeric(df['NCF? Restricted numbers'], errors='coerce').abs()
    df['Week Approved Increase'] = pd.to_numeric(df['Week Approved Increase'], errors='coerce').clip(lower=0)

    first_col = 'Project'
    type_col = 'Broad type'

    grouped = df.groupby([first_col])
    summary_df = grouped.agg({
        'Completed': 'sum',
        'Approved': 'sum',
        'Checked': 'sum',
        'Scheduled': 'sum',
        'InProgress': 'sum',
        'NCF? Restricted numbers': 'sum',
        'Week Approved Increase': 'sum'
    }).reset_index()

    summary_df['Total Completed'] = summary_df['Completed'] + summary_df['Approved'] + summary_df['Checked']
    summary_df['Total Remaining'] = summary_df['Scheduled'] + summary_df['InProgress']

    final_df = summary_df[[first_col, 'Total Completed', 'Total Remaining', 'NCF? Restricted numbers', 'Week Approved Increase']]
    final_df = final_df.rename(columns={'NCF? Restricted numbers': 'Total NCF? Restricted numbers'})

    grouped = df.groupby([type_col, first_col])
    grouped_df = grouped.agg({
        'Completed': 'sum',
        'Approved': 'sum',
        'Checked': 'sum',
        'Scheduled': 'sum',
        'InProgress': 'sum',
        'NCF? Restricted numbers': 'sum',
        'Week Approved Increase': 'sum'
    }).reset_index()

    grouped_df['Total Completed'] = grouped_df['Completed'] + grouped_df['Approved'] + grouped_df['Checked']
    grouped_df['Total Remaining'] = grouped_df['Scheduled'] + grouped_df['InProgress']

    groupedC_df = grouped_df[[type_col, first_col, 'Total Completed', 'Total Remaining', 'NCF? Restricted numbers', 'Week Approved Increase']]
    groupedC_df = groupedC_df.rename(columns={'NCF? Restricted numbers': 'Total NCF? Restricted numbers'})

    output_file = os.path.join(output_dir, f"summary_output_{D}.xlsx")

    # CREATING A PDF REPORT
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Weekly Summary Report", ln=True, align='C')

    def add_table_to_pdf(dataframe, title):
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(200, 10, txt=title.encode('latin-1', 'replace').decode('latin-1'), ln=True, align='C')
        pdf.set_font("Arial", size=10)

        col_widths = [30] * len(dataframe.columns)  # Set a fixed width per column
        row_height = 10  # Base height; will increase for multi-line
        line_spacing = 5  # Height of one line of text

        # Header
        y_before = pdf.get_y()
        x_start = pdf.get_x()
        for i, col in enumerate(dataframe.columns):
            header = str(col).encode('latin-1', 'replace').decode('latin-1')
            pdf.multi_cell(col_widths[i], line_spacing, header, border=1, align='C')
            pdf.set_xy(x_start + sum(col_widths[:i+1]), y_before)
        pdf.ln(row_height)

        # Rows
        for _, row in dataframe.iterrows():
            x_start = pdf.get_x()
            y_start = pdf.get_y()
            max_y = y_start

            # First pass to calculate max height in this row
            heights = []
            for i, item in enumerate(row):
                text = str(item).encode('latin-1', 'replace').decode('latin-1')
                lines = pdf.multi_cell(col_widths[i], line_spacing, text, border=0, align='L', split_only=True)
                heights.append(len(lines) * line_spacing)
                max_y = max(max_y, y_start + len(lines) * line_spacing)

            # Second pass to actually print the cells
            for i, item in enumerate(row):
                text = str(item).encode('latin-1', 'replace').decode('latin-1')
                pdf.set_xy(x_start + sum(col_widths[:i]), y_start)
                pdf.multi_cell(col_widths[i], line_spacing, text, border=1, align='L')

            pdf.set_y(max_y)


    # Add tables to PDF
    add_table_to_pdf(final_df, "Project Summary")
    pdf.add_page()
    add_table_to_pdf(groupedC_df, "Detailed Summary by Type")

    output_file_pdf = os.path.join(output_dir, f"New_summary_output_{D}.pdf")
    pdf.output(output_file_pdf)


    print(f"\nSummary Excel file saved to: {output_file}")

# MAIN BLOCK
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python graphgenerator.py <input_file> <output_dir>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_dir = sys.argv[2]
    generate_summary(input_file, output_dir)

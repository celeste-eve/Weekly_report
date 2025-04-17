import pandas as pd
import sys
import os
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter
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

    # WRITE TO EXCEL + CHARTS
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, sheet_name='Totals for each project', index=False)
        groupedC_df.to_excel(writer, sheet_name='Test type and project', index=False)

        workbook = writer.book
        sheet = writer.sheets['Totals for each project']

        # PIE CHART
        pie = PieChart()
        pie.title = "Weekly Completion by Project"
        labels = Reference(sheet, min_col=1, min_row=2, max_row=1 + len(final_df))
        data = Reference(sheet, min_col=5, min_row=1, max_row=1 + len(final_df))  # 'Week Approved Increase'
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(labels)
        pie.height = 7
        pie.width = 9
        sheet.add_chart(pie, "H2")

        # BAR CHART
        bar = BarChart()
        bar.type = "col"
        bar.title = "Project Progress"
        bar.y_axis.title = 'Samples'
        bar.x_axis.title = 'Project'
        bar.grouping = "stacked"

        data = Reference(sheet, min_col=2, max_col=4, min_row=1, max_row=1 + len(final_df))
        categories = Reference(sheet, min_col=1, min_row=2, max_row=1 + len(final_df))
        bar.add_data(data, titles_from_data=True)
        bar.set_categories(categories)
        bar.height = 10
        bar.width = 15
        sheet.add_chart(bar, "H20")

    print(f"\nSummary Excel file saved to: {output_file}")

# MAIN BLOCK
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python graphgenerator.py <input_file> <output_dir>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_dir = sys.argv[2]
    generate_summary(input_file, output_dir)

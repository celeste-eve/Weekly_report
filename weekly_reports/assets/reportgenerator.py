from flask import Flask, request, jsonify
import pandas as pd
import os

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files.get('file')
    output_dir = request.form.get('output_dir')

    if not file or not output_dir:
        return jsonify({'error': 'Missing file or output directory'}), 400

    df = pd.read_excel(file)

    # Clean headers if needed
    df.columns = df.columns.str.strip()

    # Extract relevant columns
    columns_needed = ['Test Type', 'Scheduled', 'InProgress', 'Completed', 'Checked', 'Approved', 'NCF? Restricted num']
    df = df.rename(columns={
        df.columns[1]: 'Project',
        df.columns[2]: 'Test Type',
        '13 Scheduled': 'Scheduled',
        'InProgress': 'InProgress',
        'Completed': 'Completed',
        'Checked': 'Checked',
        'Approved': 'Approved',
        'NCF? Restricted num': 'NCF'
    })

    summary = df.groupby(['Project', 'Test Type'], as_index=False)[['Scheduled', 'InProgress', 'Completed', 'Checked', 'Approved', 'NCF']].sum()

    output_path = os.path.join(output_dir, 'summary_output.xlsx')
    summary.to_excel(output_path, index=False)

    return jsonify({'message': 'File processed successfully', 'output_file': output_path})

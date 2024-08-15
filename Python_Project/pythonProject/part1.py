from flask import Flask, request, jsonify,send_file
import os
import pandas as pd
from reportlab.pdfgen import canvas
import zipfile
app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part in the request'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'

    file_path = os.path.join('temp', file.filename)
    file.save(file_path)

    with open(file_path, 'rb') as f:
        zip_data = zipfile.ZipFile(f)
        sheet_count = 0
        for name in zip_data.namelist():
            if name.startswith('xl/worksheets/sheet'):
                sheet_count += 1
        num_sheets = sheet_count
    route = os.path.abspath(file_path)
    return {
        'route': route,
        'num_sheets': num_sheets
    }


@app.route('/process_excel', methods=['POST'])
def process_excel():
    data = request.get_json()
    file_path = data.get('file_path')
    sheets_data = data.get('sheets_data')

    report_data = []

    try:
        excel_data = pd.read_excel(file_path, sheet_name=None)
    except Exception as e:
        return jsonify({"error": str(e)})

    for sheet_name, sheet_info in sheets_data.items():
        sheet_report = {"Sheet Name": sheet_name, "Data": {}}
        if sheet_name not in excel_data:
            sheet_report["Data"] = "Sheet not found"
        else:
            sheet_data = excel_data[sheet_name]
            data_columns = sheet_info.get("Columns", [])
            for column_name in data_columns:
                column_data = sheet_data[column_name].tolist()
                action = sheet_info.get("Action")
                if action == "Sum":
                    result = sum(column_data)
                elif action == "Average":
                    result = sum(column_data) / len(column_data)
                else:
                    result = "Invalid action"
                sheet_report["Data"][column_name] = result

        report_data.append(sheet_report)

    return jsonify(report_data)


@app.route('/generate_pdf_report', methods=['POST'])
def generate_pdf_report():
    data = request.get_json()

    pdf_filename = 'generated_report.pdf'
    c = canvas.Canvas(pdf_filename)
    c.drawString(100, 700, 'Report Data:')

    y_position = 680
    for key, value in data.items():
        c.drawString(100, y_position, f'{key}: {value}')
        y_position -= 20

    c.save()

    return send_file(pdf_filename, as_attachment=True)

if __name__ == '__main__':
    app.run()
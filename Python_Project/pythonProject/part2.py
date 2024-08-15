from flask import Flask, request
import pandas as pd
import matplotlib.pyplot as plt

app = Flask(__name__)

@app.route('/process_excel', methods=['POST'])
def process_excel():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']

    if file.filename == '':
        return "No selected file"

    if file:
        try:
            excel_data = pd.ExcelFile(file.stream)
            num_sheets = len(excel_data.sheet_names)

            total_sum = 0
            for sheet_name in excel_data.sheet_names:
                sheet_data = pd.read_excel(excel_data, sheet_name=sheet_name)
                total_sum += sheet_data.sum().sum()

            sheet_sums = {}
            for sheet_name in excel_data.sheet_names:
                sheet_data = pd.read_excel(excel_data, sheet_name=sheet_name)
                sheet_sums[sheet_name] = sheet_data.sum().sum()

            plt.bar(sheet_sums.keys(), sheet_sums.values())
            plt.xlabel('Sheet Name')
            plt.ylabel('Sum')
            plt.title('Sum of Fields for Each Sheet')
            plt.show()

            total_avg = 0
            for sheet_name in excel_data.sheet_names:
                sheet_data = pd.read_excel(excel_data, sheet_name=sheet_name)
                total_avg += sheet_data.mean().mean()
            total_avg /= num_sheets

            return f"Number of sheets: {num_sheets}, Total sum: {total_sum}, Total average: {total_avg}"
        except Exception as e:
            return str(e)

if __name__ == '__main__':
    app.run()
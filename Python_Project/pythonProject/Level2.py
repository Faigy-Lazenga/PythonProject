from flask import Flask, request, jsonify
from io import BytesIO
from reportlab.pdfgen import canvas

app = Flask(__name__)

@app.route('/generate_report', methods=['POST'])
def generate_report():
    # Extract data from the request (file name, sheet name, quantity per sheet, etc.)
    file_name = request.form.get('file_name')
    sheet_name = request.form.get('sheet_name')
    quantity_per_sheet = int(request.form.get('quantity_per_sheet'))
    average = float(request.form.get('average'))


    pdf_buffer = BytesIO()
    pdf = canvas.Canvas(pdf_buffer)
    pdf.drawString(100, 800, "File Name: " + file_name)
    pdf.drawString(100, 780, "Sheet Name: " + sheet_name)
    pdf.drawString(100, 760, "Quantity per Sheet: " + str(quantity_per_sheet))
    pdf.drawString(100, 740, "Average: " + str(average))


    pdf.save()


    response = app.make_response(pdf_buffer.getvalue())
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'attachment; filename=report.pdf'
    return response

if __name__ == '__main__':
    app.run(debug=True)
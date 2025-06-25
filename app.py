import os
from flask import Flask, request, send_file, jsonify
import pdfplumber
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def convert_pdf_to_excel(pdf_path, excel_path):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if any(cell is not None and str(cell).strip() != "" for cell in row):
                        all_rows.append(row)
    if all_rows:
        df = pd.DataFrame(all_rows)
        df.columns = all_rows[0]
        df = df[1:]
        df.to_excel(excel_path, index=False)
    else:
        raise ValueError("No valid tables found in PDF.")

@app.route('/api/convert', methods=['POST'])
def api_convert():
    if 'pdf_file' not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files['pdf_file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if not file.filename.endswith('.pdf'):
        return jsonify({"error": "Invalid file format"}), 400

    filename = secure_filename(file.filename)
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(pdf_path)

    excel_path = os.path.splitext(pdf_path)[0] + ".xlsx"

    try:
        convert_pdf_to_excel(pdf_path, excel_path)
        return send_file(excel_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

import os

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host='0.0.0.0', port=port)


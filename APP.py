from flask import Flask, request, send_file, render_template
from flask_cors import CORS
import openpyxl
import xlrd
import zipfile
import os
import tempfile
from io import BytesIO

app = Flask(__name__, template_folder="templates")
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  
CORS(app, resources={r"/*": {"origins":  ["https://gorgeous-zuccutto-a9f4d4.netlify.app"]}})

def replace_text_advanced(line, replacements):
    if line.strip() in replacements:
        return replacements[line.strip()]
    for key, val in replacements.items():
        if key in line:
            line = line.replace(key, val)
    return line

def read_excel(file):
    replacements = {}
    ext = os.path.splitext(file.filename)[1]
    if ext == '.xlsx':
        wb = openpyxl.load_workbook(file)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] and row[1]:
                replacements[str(row[0]).strip()] = str(row[1]).strip()
    elif ext == '.xls':
        book = xlrd.open_workbook(file_contents=file.read())
        sheet = book.sheet_by_index(0)
        for i in range(1, sheet.nrows):
            original, replace = sheet.row_values(i)[:2]
            if original and replace:
                replacements[str(original).strip()] = str(replace).strip()
    return replacements

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        excel_file = request.files['excel']
        dxf_files = request.files.getlist('dxfs')

        if not dxf_files:
            return "No DXF files uploaded.", 400

        replacements = read_excel(excel_file)

        with tempfile.TemporaryDirectory() as tmpdir:
            updated_paths = []
            for dxf in dxf_files:
                content = dxf.read()
                if not content:
                    continue
                lines = content.decode('latin1').splitlines()
                updated = [replace_text_advanced(line, replacements) for line in lines]

                # Only keep the filename (drop folder structure)
                file_name_only = os.path.basename(dxf.filename)
                path = os.path.join(tmpdir, file_name_only)

                with open(path, 'w', encoding='latin1', newline='') as f:
                    f.write('\r\n'.join(updated) + '\r\n')

                updated_paths.append(path)

            if not updated_paths:
                return "No valid DXF files processed.", 400

            zip_stream = BytesIO()
            with zipfile.ZipFile(zip_stream, 'w') as zipf:
                for path in updated_paths:
                    zipf.write(path, os.path.basename(path))  # add to zip without folder path

            zip_stream.seek(0)
            return send_file(zip_stream, mimetype='application/zip',
                             as_attachment=True, download_name='updated_typical.zip')
    except Exception as e:
        return f"Error: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)

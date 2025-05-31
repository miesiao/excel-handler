
from flask import Flask, request, send_from_directory, render_template, redirect, url_for
import os
from handler.order_summary import process_excel as process_summary
from handler.merge_excels import process_excel as process_merge

app = Flask(__name__)
UPLOAD_FOLDER = 'uploaded'
OUTPUT_FOLDER = 'processed'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/summary', methods=['GET', 'POST'])
def summary():
    if request.method == 'POST':
        file = request.files['file']
        if file.filename.endswith(('.xls', '.xlsx')):
            input_path = os.path.join(UPLOAD_FOLDER, file.filename)
            output_filename = f'summary_{file.filename}'
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)

            file.save(input_path)
            process_summary(input_path, output_path)

            return redirect(url_for('download', filename=output_filename))
        return "請上傳 Excel 檔案 (.xls 或 .xlsx)"
    return render_template('summary.html')

@app.route('/merge', methods=['GET', 'POST'])
def merge():
    if request.method == 'POST':
        file = request.files['file']
        if file.filename.endswith(('.xls', '.xlsx')):
            input_path = os.path.join(UPLOAD_FOLDER, file.filename)
            output_filename = f'merged_{file.filename}'
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)

            file.save(input_path)
            process_merge(input_path, output_path)

            return redirect(url_for('download', filename=output_filename))
        return "請上傳 Excel 檔案 (.xls 或 .xlsx)"
    return render_template('merge.html')

@app.route('/download/<filename>')
def download(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

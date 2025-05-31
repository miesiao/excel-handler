
from flask import Flask, request, send_from_directory, render_template, redirect, url_for
import os, uuid
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

def save_files(files, subfolder):
    folder_path = os.path.join(UPLOAD_FOLDER, subfolder)
    os.makedirs(folder_path, exist_ok=True)
    for file in files:
        if file.filename.endswith(('.xls', '.xlsx')):
            file.save(os.path.join(folder_path, file.filename))
    return folder_path

@app.route('/summary', methods=['GET', 'POST'])
def summary():
    if request.method == 'POST':
        files = request.files.getlist('files')
        if not files:
            return "請選擇至少一個 Excel 檔案"

        session_id = str(uuid.uuid4())[:8]
        folder_path = save_files(files, session_id)
        output_filename = f'summary_{session_id}.xlsx'
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        process_summary(folder_path, output_path)
        return redirect(url_for('download', filename=output_filename))
    return render_template('summary.html')

@app.route('/merge', methods=['GET', 'POST'])
def merge():
    if request.method == 'POST':
        files = request.files.getlist('files')
        if not files:
            return "請選擇至少一個 Excel 檔案"

        session_id = str(uuid.uuid4())[:8]
        folder_path = save_files(files, session_id)
        output_filename = f'merged_{session_id}.xlsx'
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        process_merge(folder_path, output_path)
        return redirect(url_for('download', filename=output_filename))
    return render_template('merge.html')

@app.route('/download/<filename>')
def download(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

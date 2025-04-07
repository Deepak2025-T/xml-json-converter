from flask import Flask, render_template, request, redirect, send_file, flash
import os
from werkzeug.utils import secure_filename
from converter import json_to_excel, xml_to_excel, fetch_and_convert, database

app = Flask(__name__)
app.secret_key = 'supersecretkey'
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

database.initialize_db()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload/xml', methods=['GET', 'POST'])
def upload_xml():
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith('.xml'):
            filename = secure_filename(file.filename)
            path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(path)
            out_path = xml_to_excel.convert_xml_to_excel(path, CONVERTED_FOLDER)
            database.save_file_record(filename, 'xml', 'upload', out_path)
            return send_file(out_path, as_attachment=True)
        flash("Please upload a valid XML file.")
    return render_template('xml_upload.html')

@app.route('/upload/json', methods=['GET', 'POST'])
def upload_json():
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith('.json'):
            filename = secure_filename(file.filename)
            path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(path)
            out_path = json_to_excel.convert_json_to_excel(path, CONVERTED_FOLDER)
            database.save_file_record(filename, 'json', 'upload', out_path)
            return send_file(out_path, as_attachment=True)
        flash("Please upload a valid JSON file.")
    return render_template('json_upload.html')

@app.route('/fetch', methods=['GET', 'POST'])
def fetch_url():
    if request.method == 'POST':
        url = request.form['url']
        out_path, file_type, filename = fetch_and_convert.fetch_file_from_url(url, CONVERTED_FOLDER)
        database.save_file_record(filename, file_type, 'url', out_path, url)
        return send_file(out_path, as_attachment=True)
    return render_template('fetch_url.html')

if __name__ == '__main__':
    app.run(debug=True)
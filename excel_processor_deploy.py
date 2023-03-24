import os
from flask import Flask, request, send_from_directory
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def process_xlsx(file_name):
    # Add your existing processing code here

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename) # type: ignore
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            process_xlsx(file_path)
            return send_from_directory(app.config['UPLOAD_FOLDER'], 'processed_Excel_file.xlsx', as_attachment=True)
        else:
            return 'Invalid file format. Please upload an .xlsx file.', 400
    return '''
    <!doctype html>
    <title>Upload an .xlsx File</title>
    <h1>Upload an .xlsx File</h1>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file>
      <input type=submit value=Upload>
    </form>
    '''

if __name__ == '__main__':
    app.run(debug=True)

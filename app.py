from flask import Flask, request, jsonify, send_file, make_response
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import re
import docx2txt
from PyPDF2 import PdfReader
import xlwt

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf', 'docx'}
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, 'rb') as file:
        pdf_reader = PdfReader(file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
    return text


def extract_text_from_docx(docx_path):
    text = docx2txt.process(docx_path)
    return text


def extract_emails(text):
    emails = re.findall(r'[\w\.-]+@[\w\.-]+', text) or re.findall(r'[789]\d{9}$',text)
    return emails


def extract_phone_numbers(text):
    phone_numbers = re.findall(r'(\+\d{1,2}\s?)?(\d{3}\s?\d{3}\s?\d{4})', text)
    return ["".join(pn) for pn in phone_numbers]


@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files[]' not in request.files:
        return jsonify({'error': 'No file part'})

    files = request.files.getlist('files[]')
        
    output_workbook = xlwt.Workbook()
    output_sheet = output_workbook.add_sheet('Extracted Data')

    row = 0
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            if filename.endswith('.pdf'):
                text = extract_text_from_pdf(file_path)
            elif filename.endswith('.docx'):
                text = extract_text_from_docx(file_path)
            else:
                continue

            emails = extract_emails(text)
            phone_numbers = extract_phone_numbers(text)
            output_sheet.write(row, 0, filename)
            for i in range(len(emails)):
                output_sheet.write(row + i, 1, emails[i])
                

            for j in range(len(phone_numbers)):
                output_sheet.write(row + j, 2, phone_numbers[j])
            
            
            output_sheet.write(row,3, text)
            row +=( len(emails) if len(emails) > len(phone_numbers) else len(phone_numbers) ) + 1

    output_workbook.save('extracted_data.xls')
    return jsonify({'success': True})


@app.route('/download', methods=['GET'])
def download_file():
    # Secure the filename
    filename = 'extracted_data.xls'
    # Check if file exists
    file_path = os.path.join(filename)
    if not os.path.exists(file_path):
        return 'File not found!', 404  # Return a 404 Not Found error

    # Read file data (optional, for processing)
    with open(file_path, 'rb') as f:
        file_data = f.read()

    # Set response headers for download
    response = make_response(file_data)
    # Excel MIME type
    response.headers['Content-Type'] = 'application/vnd.ms-excel'
    response.headers['Content-Disposition'] = f'attachment; filename="{filename}"'
    
    return response

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0')

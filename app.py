import os
from flask import Flask, render_template, request, send_file
from openpyxl import Workbook
import tempfile
import re
from PyPDF2 import PdfReader
from docx import Document

app = Flask(__name__, static_url_path='/static')

def extract_information_from_file(file_path):
    extracted_text = ''
    emails = []
    phone_numbers = []

    try:
        if file_path.endswith('.pdf'):
            pdf_file = open(file_path, 'rb')
            pdf_reader = PdfReader(pdf_file)
            for page in pdf_reader.pages:
                extracted_text += page.extract_text()
            pdf_file.close()

        elif file_path.endswith('.docx') or file_path.endswith('.doc'):
            docx_file = Document(file_path)
            for paragraph in docx_file.paragraphs:
                extracted_text += paragraph.text + '\n'

        # Extract emails and phone numbers using regex
        emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", extracted_text)
        phone_numbers = re.findall(r'[\+\(]?[1-9][0-9]{8,}[0-9]', extracted_text)

    except Exception as e:
        print(f"Error extracting information from {file_path}: {e}")

    return emails, phone_numbers, extracted_text

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        files = request.files.getlist('files[]')  # Access uploaded files as a list

        if len(files) == 0:
            return "No files selected"

        folder_path = tempfile.mkdtemp()

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Extracted Information"
            ws.append(["File Name", "Email", "Phone Number", "Text"])

            for file in files:
                if file.filename == '':
                    continue

                file_path = os.path.join(folder_path, file.filename)
                file.save(file_path)

                emails, phone_numbers, text = extract_information_from_file(file_path)

                ws.append([file.filename, ", ".join(emails), ", ".join(phone_numbers), text])

            output_filename = "uploaded_files_information.xlsx"
            wb.save(output_filename)

            return send_file(output_filename, as_attachment=True)

        except Exception as e:
            return f"An error occurred: {e}"

        finally:
            for filename in os.listdir(folder_path):
                file_path = os.path.join(folder_path, filename)
                os.remove(file_path)
            os.rmdir(folder_path)

    return render_template('index.html')

# Remove the app.run() part as it will be handled by GitHub Actions
if __name__ == '__main__':
    app.run(debug=False,host='0.0.0.0',port=8080)


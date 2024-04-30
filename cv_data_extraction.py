import os
import re
import pandas as pd
import PyPDF2
import docx
import fitz
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        num_pages = len(pdf_reader.pages)
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
    return text

# Other extract functions remain unchanged...

def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        try:
            cv_data = []
            files = request.files.getlist('file[]')
            for file in files:
                if file.filename != '':
                    # Validate file extension
                    if file.filename.split('.')[-1] not in ['pdf', 'docx', 'doc']:
                        return "Error: Unsupported file format"
                    cv_info = extract_cv_data(file)
                    cv_info['File'] = file.filename
                    cv_data.append(cv_info)
            output_file = 'cv_data.xlsx'
            save_to_excel(cv_data, output_file)
            return send_file(output_file, as_attachment=True)
        except FileNotFoundError:
            return "Error: File not found"
        except PermissionError:
            return "Error: Permission denied"
        except Exception as e:
            return f"An error occurred: {str(e)}"
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)

import os
import PyPDF2
import re
import pandas as pd
import docx
import fitz

def extract_text_from_pdf(pdf_path):
    text = ""
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        num_pages = len(pdf_reader.pages)
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
    return text

def extract_text_from_docx(docx_path):
    doc = docx.Document(docx_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def extract_text_from_doc(doc_path):
    doc = fitz.open(doc_path)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def extract_text_from_file(file_path):
    if file_path.endswith('.pdf'):
        return extract_text_from_pdf(file_path)
    elif file_path.endswith('.docx'):
        return extract_text_from_docx(file_path)
    elif file_path.endswith('.doc'):
        return extract_text_from_doc(file_path)
    else:
        raise ValueError("Unsupported file format")

def extract_contact_info(text):
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_pattern = r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})'

    emails = re.findall(email_pattern, text)
    phones = re.findall(phone_pattern, text)

    return emails, phones

def extract_cv_data(file_path):
    text = extract_text_from_file(file_path)
    emails, phones = extract_contact_info(text)
    return {'Email': emails, 'Phone': phones, 'Text': text}

def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)

def process_cv_folder(input_folder, output_file):
    cv_data = []
    for filename in os.listdir(input_folder):
        if filename.endswith('.pdf') or filename.endswith('.docx') or filename.endswith('.doc'):
            cv_path = os.path.join(input_folder, filename)
            cv_info = extract_cv_data(cv_path)
            cv_info['File'] = filename
            cv_data.append(cv_info)
    save_to_excel(cv_data, output_file)

input_folder = r'C:\Users\shrey\OneDrive\Desktop\Sample2'   
output_file = 'cv_data.xlsx' 
process_cv_folder(input_folder, output_file)

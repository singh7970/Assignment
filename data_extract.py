import docx2txt
import re
import pandas as pd
from PyPDF2 import PdfReader

def extract_info_from_cv(cv_text):
    email_pattern = r'[\w\.-]+@[\w\.-]+'
    phone_pattern = r'[6789]*[0-9]{9}'
    
    emails = re.findall(email_pattern, cv_text)
    phones = re.findall(phone_pattern, cv_text)
    
    return emails, phones, cv_text

def extract_text_from_pdf(pdf_file):
    text = ""
    with open(pdf_file, "rb") as file:
        reader = PdfReader(file)
        for page_number in range(len(reader.pages)):
            text += reader.pages[page_number].extract_text()
    return text

def save_to_excel(data, output_file):
    df = pd.DataFrame(data, columns=['Email', 'Phone', 'Text'])
    df.to_excel(output_file, index=False)

def main(cv_files, pdf_files, output_file):
    data = []
    for cv_file in cv_files:
        cv_text = docx2txt.process(cv_file)
        emails, phones, text = extract_info_from_cv(cv_text)
        data.append((emails, phones, text))
    
    for pdf_file in pdf_files:
        pdf_text = extract_text_from_pdf(pdf_file)
        emails, phones, text = extract_info_from_cv(pdf_text)
        data.append((emails, phones, text))
    
    save_to_excel(data, output_file)
    print("Data extracted and saved to", output_file)

if __name__ == "__main__":
    cv_files = ["AkashGoel.docx"]  # Add your CV file names here
    pdf_files = ["CAChamanKumar.pdf"]  # Add your PDF file names here
    output_file = "cv_output_data.xlsx"
    main(cv_files, pdf_files, output_file)

# pip install openpyxl
# pip install PyPDF2 pandas
import os
import re
import PyPDF2
import pandas as pd

def extract_name_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfFileReader(file)
        text = ""
        for page_num in range(pdf_reader.numPages):
            text += pdf_reader.getPage(page_num).extractText()
    return text
def extract_name_and_roll_from_filename(filename):
    match = re.match(r'(\w+)_(\w+)', filename)
    if match:
        return match.group(2), match.group(1)
    else:
        return None, None
def process_pdfs(folder_path, output_excel_path):
    data = {'File Name': [], 'Roll': [], 'Name': []}

    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            roll, name = extract_name_and_roll_from_filename(filename)
            if roll is None or name is None:
                print(f"Unable to extract roll and name from filename: {filename}")
            else:
                data['File Name'].append(filename)
                data['Roll'].append(roll)
                data['Name'].append(name)

    df = pd.DataFrame(data)
    df.to_excel(output_excel_path, index=False)
    print(f"Data has been extracted and saved to {output_excel_path}")

if __name__ == "__main__":
    folder_path = r"C:\Users\dfgpi\Music\Course"
    output_excel_path = r"C:\Users\dfgpi\Music\MongoDbList.xlsx"
    process_pdfs(folder_path, output_excel_path)

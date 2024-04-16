import pytesseract
from PIL import Image
import PyPDF2
import pdf2image
import re
import os
import dotenv

dotenv.load_dotenv()

pytesseract.pytesseract.tesseract_cmd = os.environ.get('TESSERACT_CMD')

custom_config = r'--oem 3 --psm 6 -l pol'

# print(pytesseract)
file_path = 'a.pdf'
poppler_path=os.environ.get('POPPLER_PATH')

images = pdf2image.convert_from_path(file_path, poppler_path=poppler_path)

def perform_ocr(images):
    extracted_text = []
    for image in images:
        text = pytesseract.image_to_string(image, config=custom_config)
        extracted_text.append(text)
    return '\n'.join(extracted_text)

ocr_text = perform_ocr(images)
# print(ocr_text)

name_pattern = r'(?<=na rzecz )[A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż]+\s[A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż]+\b'
name = re.findall(name_pattern, ocr_text)
name = name[0]

vin_pattern = r'(?<=VIN: )[A-Z0-9]{4,17}\b'
vin = re.findall(vin_pattern, ocr_text)

tr_pattern = r'(?<=nr rej. )([A-Z0-9]+\s[A-Z0-9]+)\b'
tr = re.findall(tr_pattern, ocr_text)

print(name)
print(vin)
print(tr[0].replace('\n', ' '))
sentences = [name, vin[0], tr[0].replace('\n', ' ')]

from openpyxl import Workbook

def write_to_excel_from_ocr(sentences, excel_file_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Extracted Sentences (OCR)'

    # for idx, sentence in enumerate(sentences, start=1):
    #     sheet[f'A{idx}'] = sentence
    sheet[f'A1'] = sentences[0]
    sheet[f'B1'] = sentences[1]
    sheet[f'C1'] = sentences[2]
    workbook.save(excel_file_path)

# Usage:
write_to_excel_from_ocr(sentences, 'output_ocr.xlsx')

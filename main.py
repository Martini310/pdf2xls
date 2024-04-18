import pytesseract
from PIL import Image
import PyPDF2
import pdf2image
import re
import os
import dotenv
from openpyxl import Workbook
import docx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


dotenv.load_dotenv()

pytesseract.pytesseract.tesseract_cmd = os.environ.get('TESSERACT_CMD')

custom_config = r'--oem 3 --psm 6 -l pol'
poppler_path=os.environ.get('POPPLER_PATH')

file_path = 'a.pdf'
files = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.pdf')]
print(files)
# images = pdf2image.convert_from_path(file_path, poppler_path=poppler_path)
images = [pdf2image.convert_from_path(f, poppler_path=poppler_path) for f in files]

def perform_ocr(images):
    ocr_text = []
    for image in images:
        extracted_text = []
        for page in image:
            text = pytesseract.image_to_string(page, config=custom_config)
            extracted_text.append(text)
        ocr_text.append('\n'.join(extracted_text))
    return ocr_text

ocr_text = perform_ocr(images)
# print(ocr_text)

def find_patterns(ocr_text):
    output = {}
    for text in ocr_text:
        try:
            kt_pattern = r'(?<=KT.5410.[0-9].)([0-9]+)'
            kt = re.findall(kt_pattern, text)
            kt = kt[0]
            
            # print(text)
            name_pattern = r'(?<=na rzecz )(([A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż]+\s*)+)'
            name = re.search(name_pattern, text)
            # print(name)
            name = name[0].replace('\n', ' ')

            vin_pattern = r'(?<=VIN:)[\s —-]*([A-Z0-9]*)'
            vin = re.findall(vin_pattern, text)

            tr_pattern = r'(?<=nr rej\.)\s?([A-Z0-9]+\s[A-Z0-9]+)\b'
            tr = re.findall(tr_pattern, text)
            tr = tr[0].replace('\n', ' ')
            
            art_pattern = r'(?<=w związku z art\. ).*(?= ustawy)'
            art = re.search(art_pattern, text)
            
            address_pattern = r'(?<=na rzecz )([\w\s©\[\],]+)(?=w związku)'
            address = re.search(address_pattern, text)
            
            date_pattern = r'(?<=Poznań, dnia ).+(?=r)'
            date = re.search(date_pattern, text)
            
            # brand_pattern = r'(?<=marki).*(?=o nr rej)'
            # brand = re.findall(brand_pattern, text)
            # brand = brand[0]
            
            print(kt, name, vin[0], tr, art[0], address[0], date[0])
            output[kt] = {'name': name, 'vin': vin[0], 'tr': tr}
        except IndexError as e:
            # print(text)
            print(kt, 'error', e)
            print(kt, name, vin, tr, art, address, date)
            print('---' * 50)
        except TypeError as e:
            print(kt, 'type error')
            print(text)
    return output


sentences = find_patterns(ocr_text)
# print(sentences)


def write_to_excel_from_ocr(sentences, excel_file_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Extracted Sentences (OCR)'

    for idx, key in enumerate(sentences, start=1):
        # sheet[f'A{idx}'] = sentence
        sheet[f'A{idx}'] = key
        sheet[f'B{idx}'] = sentences[key]['name']
        sheet[f'C{idx}'] = sentences[key]['vin']
        sheet[f'D{idx}'] = sentences[key]['tr']
    workbook.save(excel_file_path)

# Usage:
write_to_excel_from_ocr(sentences, 'output_ocr.xlsx')

doc = Document()

font = doc.styles['Normal'].font

font.name = 'Calibri'
font.size = Pt(10)

paragraph = doc.add_paragraph('Starostwo Powiatowe w Poznaniu\nul. Jackowskiego 18\n60-509 Poznań')
paragraph_format = paragraph.paragraph_format

paragraph_format.line_spacing = 0.75


paragraph = doc.add_paragraph('Adam Nowak\nul. Długa 1\n12-345 Zbąszyń')
paragraph_format = paragraph.paragraph_format

paragraph_format.line_spacing = 0.75
paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT

doc.save('test.docx')
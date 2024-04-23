import os
import re
from typing import List, Dict
from PIL import Image
import pdf2image
import dotenv
from openpyxl import Workbook
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pytesseract


dotenv.load_dotenv()

pytesseract.pytesseract.tesseract_cmd = os.environ.get('TESSERACT_CMD')

CUSTOM_CONFIG = r'--oem 3 --psm 6 -l pol'
poppler_path=os.environ.get('POPPLER_PATH')

pdf_files = [os.path.join("./skany", f) for f in os.listdir('./skany') if os.path.isfile(os.path.join("./skany", f)) and f.endswith('.pdf')]
# images = [pdf2image.convert_from_path(f, poppler_path=poppler_path) for f in list(reversed(pdf_files))]



class ReadPDF:
    def __init__(self, path: str, reverse: bool = False) -> None:
        self.path: str = path
        self.reverse: bool = reverse
        self.files: List[str] = [
            os.path.join(self.path, f)
            for f in os.listdir(self.path)
            if os.path.isfile(os.path.join(self.path, f)) and f.endswith('.pdf')
        ]

    def create_images(self, files: List[str]) -> List[List[Image.Image]]:
        """
        Convert PDF files to a list of images.
        """
        images: List[List[Image.Image]] = [
            pdf2image.convert_from_path(f, poppler_path=poppler_path)
            for f in files
        ]
        if self.reverse:
            images.reverse()
        return images

    def perform_ocr(self, images: List[List[Image.Image]]) -> List[str]:
        """
        Perform OCR (Optical Character Recognition) on a list of images.
        """
        text: List[str] = []
        for image in images:
            extracted_text = []
            for page in image:
                text = pytesseract.image_to_string(page, config=CUSTOM_CONFIG)
                extracted_text.append(text)
            text.append('\n'.join(extracted_text))
        return text

    def find_patterns(self, ocr_text: List[str]) -> Dict[str, Dict[str, str]]:
        """
        Find specific patterns in OCR text and extract relevant information.
        """
        output: Dict[str, Dict[str, str]] = {}
        kt = ''
        for text in ocr_text:
            try:
                tmp = kt if kt else '1'
                kt_pattern = r'(?<=KT.5410.[0-9].)([0-9]+)'
                kt = re.findall(kt_pattern, text)
                if not kt:
                    kt = str(int(tmp) + 1)
                else:
                    kt = kt[0]

                # print(text)
                name_pattern = r'(?<=na rzecz )(([A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż]+\s*)+)'
                client_name = re.search(name_pattern, text)
                # print(name)
                client_name = client_name[0].replace('\n', ' ').strip()

                vin_pattern = r'(?<=VIN:)[\s —-]*([A-Z0-9—-]*)'
                vin = re.findall(vin_pattern, text)
                vin = 'błąd odczytu' if not vin else vin[0]

                art_pattern = r'(?<=w związku z art\. )[\w\s\.]*(?= ustawy)'
                art = re.search(art_pattern, text)
                art = art[0].replace('\n', ' ')

                if '73aa ust. 1 pkt 3' in art:
                    tr = ''
                else:
                    tr_pattern = r'(?<=rej\.)\s*([A-Z0-9]+\s*[A-Z0-9]+)\b'
                    tr = re.findall(tr_pattern, text)
                    print(tr)
                    tr = tr[0].replace('\n', ' ')

                czynnosc = ''
                if '73aa ust. 1 pkt 3' in art:
                    czynnosc = 'SPROWADZONY'
                elif '73aa ust. 1 pkt 1' in art:
                    czynnosc = 'NABYCIE'
                elif '78 ust. 2 pkt 1' in art:
                    czynnosc = 'ZBYCIE'

                address_pattern = r'(?<=na rzecz )[\s\w,.©\[\]/\\-]*(?=w związku)'
                address = re.search(address_pattern, text)
                address = 'błąd odczytu' if not address else address[0]

                date_pattern = r'(?<=Poznań, dnia ).+(?=r)'
                date = re.search(date_pattern, text)
                date = date[0].replace('—', '.').replace('-', '.')

                brand_pattern = r'(?<=marki\s)[\w\s\\/-]+(?=o)'
                brand = re.findall(brand_pattern, text)
                brand = brand[0].strip()

                output[kt] = {
                    'name': client_name,
                    'vin': vin[0],
                    'tr': tr,
                    'date': date,
                    'art': art,
                    'czynnosc': czynnosc
                }
                
            except IndexError as e:
                print(text)
                print(kt, 'error', e)
                print(kt, client_name, vin, tr, art, date, brand, address)
                print('---' * 50)
            except TypeError as e:
                print(kt, 'type error', e)
                print(text)
        return output

    def write_to_excel(self, data: Dict[str, Dict[str, str]], excel_file_path: str) -> None:
        """
        Write data extracted from OCR to an Excel file.
        """
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Extracted data (OCR)'

        for idx, kt in enumerate(data, start=1):
            sheet[f'A{idx}'] = kt
            sheet[f'B{idx}'] = data[kt]['tr']
            sheet[f'C{idx}'] = data[kt]['vin']
            sheet[f'D{idx}'] = data[kt]['name']
            sheet[f'E{idx}'] = ''
            sheet[f'F{idx}'] = ''
            sheet[f'G{idx}'] = data[kt]['date']
            sheet[f'H{idx}'] = ''
            sheet[f'I{idx}'] = ''
            sheet[f'J{idx}'] = ''
            sheet[f'K{idx}'] = data[kt]['czynnosc']
            sheet[f'L{idx}'] = data[kt]['art']
        workbook.save(excel_file_path)


def perform_ocr(images):
    ocr_text = []
    for image in images:
        extracted_text = []
        for page in image:
            text = pytesseract.image_to_string(page, config=CUSTOM_CONFIG)
            extracted_text.append(text)
        ocr_text.append('\n'.join(extracted_text))
    return ocr_text

# ocr_text = perform_ocr(images)
# print(ocr_text)

def find_patterns(ocr_text):
    output = {}
    kt = ''
    for text in ocr_text:
        try:
            tmp = kt if kt else '0'
            kt_pattern = r'(?<=KT.5410.[0-9].)([0-9]+)'
            kt = re.findall(kt_pattern, text)
            if not kt:
                kt = str(int(tmp) + 1)
            else:
                kt = kt[0]
            
            # print(text)
            name_pattern = r'(?<=na rzecz )(([A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż]+\s*)+)'
            name = re.search(name_pattern, text)
            # print(name)
            name = name[0].replace('\n', ' ').strip()

            vin_pattern = r'(?<=VIN:)[\s —-]*([A-Z0-9—-]*)'
            vin = re.findall(vin_pattern, text)
            vin = 'błąd odczytu' if not vin else vin

            art_pattern = r'(?<=w związku z art\. )[\w\s\.]*(?= ustawy)'
            art = re.search(art_pattern, text)
            art = art[0].replace('\n', ' ')

            if '73aa ust. 1 pkt 3' in art:
                tr = ''
            else:
                tr_pattern = r'(?<=rej\.)\s*([A-Z0-9]+\s*[A-Z0-9]+)\b'
                tr = re.findall(tr_pattern, text)
                print(tr)
                tr = tr[0].replace('\n', ' ')

            czynnosc = ''
            if '73aa ust. 1 pkt 3' in art:
                czynnosc = 'SPROWADZONY'
            elif '73aa ust. 1 pkt 1' in art:
                czynnosc = 'NABYCIE'
            elif '78 ust. 2 pkt 1' in art:
                czynnosc = 'ZBYCIE'

            address_pattern = r'(?<=na rzecz )[\s\w,.©\[\]/\\-]*(?=w związku)'
            address = re.search(address_pattern, text)
            address = 'błąd odczytu' if not address else address[0]

            date_pattern = r'(?<=Poznań, dnia ).+(?=r)'
            date = re.search(date_pattern, text)
            date = date[0].replace('—', '.').replace('-', '.')
            
            brand_pattern = r'(?<=marki\s)[\w\s\\/-]+(?=o)'
            brand = re.findall(brand_pattern, text)
            brand = brand[0].strip()
            
            # print(kt, name, vin[0], tr, art[0], address[0], date[0])
            output[kt] = {'name': name, 'vin': vin[0], 'tr': tr, 'date': date, 'art': art, 'czynnosc': czynnosc}
        except IndexError as e:
            print(text)
            print(kt, 'error', e)
            print(kt, name, vin, tr, art, date, brand, address)
            print('---' * 50)
        except TypeError as e:
            print(kt, 'type error', e)
            print(text)
    return output


# sentences = find_patterns(ocr_text)
# print(sentences)


def write_to_excel_from_ocr(sentences, excel_file_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Extracted Sentences (OCR)'

    for idx, key in enumerate(sentences, start=1):
        sheet[f'A{idx}'] = key
        sheet[f'B{idx}'] = sentences[key]['tr']
        sheet[f'C{idx}'] = sentences[key]['vin']
        sheet[f'D{idx}'] = sentences[key]['name']
        sheet[f'E{idx}'] = ''
        sheet[f'F{idx}'] = ''
        sheet[f'G{idx}'] = sentences[key]['date']
        sheet[f'H{idx}'] = ''
        sheet[f'I{idx}'] = ''
        sheet[f'J{idx}'] = ''
        sheet[f'K{idx}'] = sentences[key]['czynnosc']
        sheet[f'L{idx}'] = sentences[key]['art']
    workbook.save(excel_file_path)

# Usage:
# write_to_excel_from_ocr(sentences, 'output_ocr.xlsx')


def create_docx():
    doc = Document()

    font = doc.styles['Normal'].font

    font.name = 'Calibri'
    font.size = Pt(10)

    paragraph = doc.add_paragraph('Poznań dnia \n')
    paragraph_format = paragraph.paragraph_format

    paragraph_format.line_spacing = 0.75
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT


    paragraph = doc.add_paragraph('Starostwo Powiatowe w Poznaniu\nul. Jackowskiego 18\n60-509 Poznań')
    paragraph_format = paragraph.paragraph_format

    paragraph_format.line_spacing = 0.75
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT


    paragraph = doc.add_paragraph('\nDECYZJA NR ')
    paragraph_format = paragraph.paragraph_format

    paragraph_format.line_spacing = 0.75
    paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.runs[0].bold = True


    name = 'Martin Brzeziński'
    pesel = '293793759'
    marka = 'ford'
    tr = 'PZ 12345'
    vin = 'JFKKSAFHYJK98987'

    podstawa_prawna = f"""Na podstawie art. 104 ustawy z dnia 14 czerwca 1960 r. – kodeks postepowania administracyjnego (Dz. U. z 2023 r., poz. 775 t. j.) w związku z art. 140mb ust. 1 oraz art. 73aa ust. 1 pkt 1 ustawy z dnia 20 czerwca 1997 r. – Prawo o ruchu drogowym (Dz. U. z 2023 r., poz. 1047 t. j.) po rozpatrzeniu sprawy Pana/i {name}. (PESEL: {pesel}) będącego właścicielem pojazdu marki {marka}. o numerze rejestracyjnym {tr}, nr nadwozia: {vin}."""

    paragraph = doc.add_paragraph(podstawa_prawna)
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


    paragraph = doc.add_paragraph('\nSTAROSTA\n')
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.runs[0].bold = True

    kara = """nakłada karę pieniężną w wysokości 500 zł (słownie: pięćset zł) w związku z niedopełnieniem obowiązku złożenia wniosku o rejestrację w terminie 30 dni od dnia nabycia wyżej wymienionego pojazdu, tj. obowiązku wynikającego z art. 73aa ust. 1 pkt 1 ustawy - Prawo o ruchu drogowym."""

    paragraph = doc.add_paragraph(kara)
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.save('test.docx')

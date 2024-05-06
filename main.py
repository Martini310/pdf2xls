import os
import re
from typing import List, Dict, Pattern
from PIL import Image
import pdf2image
import dotenv
from openpyxl import Workbook
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
import pytesseract
from datetime import date
import json

dotenv.load_dotenv()

pytesseract.pytesseract.tesseract_cmd = os.environ.get('TESSERACT_CMD')

CUSTOM_CONFIG = r'--oem 3 --psm 6 -l pol'
poppler_path=os.environ.get('POPPLER_PATH')

pdf_files = [os.path.join("./skany", f) for f in os.listdir('./skany') if os.path.isfile(os.path.join("./skany", f)) and f.endswith('.pdf')]
# images = [pdf2image.convert_from_path(f, poppler_path=poppler_path) for f in list(reversed(pdf_files))]


class PDFHandler:
    patterns: Dict[str, Pattern[str]] = {
        'kt': r'(?<=KT.5410.[0-9].)([0-9]+)',
        'name': r'(?<=na rzecz )([A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż]+(?:\s+[A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż]+)*(?:\s+[A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż]+))', # r'(?<=na rzecz )(([A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż]+\s*)+)',
        'vin': r'(?<=VIN:)[\s —-]*([A-Z0-9—-]*)',
        'basis': r'(?<=w związku z art\. )[\w\s\.]*(?= ustawy)',
        'tr': r'(?<=rej\.)\s*([A-Z0-9]+\s*[A-Z0-9]+)\b',
        'address': r'(?<=na rzecz )[\s\w,.©\[\]/\\-]*(?=w związku)',
        'date': r'(?<=Poznań, dnia ).+(?=r)',
        'brand': r'(?<=marki\s)[\w\s\\/-]+(?=o)',
        'pesel': r'[0-9]{9,11}',
        'purchase_date': r'(?<=z dnia )[0-9-.]+(?=r.)',
    }

    def __init__(self, path: str) -> None:
        self.path: str = path
        self.text = self.perform_ocr(self.create_images(self.path))
        self.results = self.extract_text(self.text, self.patterns)
        self.przypisz_czynnosc()

    def create_images(self, file: str) -> List[Image.Image]:
        """
        Convert PDF file into image.
        """
        image: List[Image.Image] = pdf2image.convert_from_path(file, poppler_path=poppler_path)

        return image

    def perform_ocr(self, image: List[Image.Image]) -> str:
        """
        Perform OCR (Optical Character Recognition) on an image.
        """
        extracted_text = []
        for page in image:
            text = pytesseract.image_to_string(page, config=CUSTOM_CONFIG)
            extracted_text.append(text)

        return '\n'.join(extracted_text)

    def find_pattern(self, text: str, pattern: Pattern[str]) -> str:
        """
        Find and return provided pattern in a given text
        """
        try:
            matches: List[str] = re.findall(pattern, text)
            if not matches:
                return 'n/d'
            result = matches[0]
            return result
        except IndexError as e:
            print(text)
            print('error', e)
            print('---' * 50)
            return 'błąd'
        except TypeError as e:
            print('type error', e)
            print(text)
            return 'błąd'

    def extract_text(self, text: str, patterns: Dict[str, Pattern[str]]) -> Dict[str, str]:
        """
        Return a dict with extracted data from given text based on patterns provided in dict
        """
        data: Dict[str, str] = {}
        for key, pattern in patterns.items():
            extracted_text = self.find_pattern(text, pattern)
            extracted_text = extracted_text.strip().replace('\n', ' ')
            data[key] = extracted_text
        return data
    
    def przypisz_czynnosc(self) -> None:
        czynnosc = 'n/d'
        if '73aa ust. 1 pkt 3' in self.results['basis']:
            czynnosc = 'SPROWADZONY'
        elif '73aa ust. 1 pkt 1' in self.results['basis']:
            czynnosc = 'NABYCIE'
        elif '78 ust. 2 pkt 1' in self.results['basis']:
            czynnosc = 'ZBYCIE'
        self.results['czynność'] = czynnosc

    def create_docx(self) -> None:
        data = self.results

        with open('docx_source_text.json', 'r', encoding='utf-8') as file:
            source = json.load(file)

        doc = Document()

        # Set custom style for bold centered titles
        styles = doc.styles
        title = styles.add_style('Tytuł', WD_STYLE_TYPE.PARAGRAPH)
        title.base_style = styles['Normal']
        title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

        font = doc.styles['Tytuł'].font
        font.bold = True

        # Customize base style
        doc.styles['Normal'].paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        doc.styles['Normal'].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        doc.styles['Normal'].paragraph_format.space_after = Cm(0)

        font = doc.styles['Normal'].font
        font.name = 'Calibri'
        font.size = Pt(10)

        # Customize page and margins sizes
        section = doc.sections[0]

        section.page_width = Inches(8.27)   # Equivalent to 210 mm
        section.page_height = Inches(11.69)  # Equivalent to 297 mm

        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

        # Choose right template
        if self.results['czynność'] == 'NABYCIE':
            kara = source['kara_nabycie']
            uzasadnienie = source['uzasadnienie_nabycie']
        elif self.results['czynność'] == 'SPROWADZONY':
            kara = source['kara_ue']
            uzasadnienie = source['uzasadnienie_ue']
        elif self.results['czynność'] == 'ZBYCIE':
            pass
        
        uzasadnienie += source['uzasadnienie_wspolne']

        doc.add_paragraph(f'Poznań dnia {date.today()}').paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        header = 'Starosta Poznański\nul. Jackowskiego 18\n60-509 Poznań'
        doc.add_paragraph(header).paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

        paragraph = doc.add_paragraph(f"\n\nDECYZJA NR KT.5410.7.{data['kt']}.2024\n", style=title)

        doc.add_paragraph(source['podstawa_prawna'].format(data['name'], data['pesel'], data['brand'], data['tr'], data['vin']))

        paragraph = doc.add_paragraph('\nStarosta\n', style=title)
        paragraph = doc.add_paragraph(kara)
        paragraph = doc.add_paragraph('\nUzasadnienie\n', style=title)

        doc.add_paragraph(uzasadnienie[0].format(data['name'], data['purchase_date']))
        doc.add_paragraph(uzasadnienie[1])
        doc.add_paragraph(uzasadnienie[2].format(data['date']))
        for uzasadnienie in uzasadnienie[3:]:
            doc.add_paragraph(uzasadnienie)

        add_numbered_paragraphs(doc, source['przepisy'][:7], 'ListNumber', Inches(0.5))

        for przepis in source['przepisy'][7:10]:
            doc.add_paragraph(przepis)

        add_numbered_paragraphs(doc, source['przepisy'][10:12], 'List Number 2', space_after=Cm(0))
        paragraph = doc.add_paragraph(source['przepisy'][12])

        for przepis in source['przepisy'][13:15]:
            paragraph = doc.add_paragraph(przepis)
            paragraph.paragraph_format.left_indent = Inches(0.25)

        for przepis in source['przepisy'][15:]:
            paragraph = doc.add_paragraph(przepis)

        paragraph = doc.add_paragraph('\nPouczenie\n', style=title)
        paragraph = doc.add_paragraph(source['pouczenia'][0])
        paragraph = doc.add_paragraph('\n')

        paragraph.add_run('\tWpłaty należy dokonać na konto numer: ')
        paragraph.add_run('7710 3012 4700 0000 0034 9162 41').bold = True
        paragraph.add_run(' w tytule podając nr decyzji ')
        paragraph.add_run(f"KT.5410.7.{data['kt']}.2024").bold = True
        paragraph = doc.add_paragraph()
        
        paragraph = doc.add_paragraph(source['pouczenia'][2])
        paragraph = doc.add_paragraph()
        paragraph = doc.add_paragraph(source['pouczenia'][3])
        paragraph = doc.add_paragraph('\n' * 9)
        
        paragraph = doc.add_paragraph('Otrzymują:')

        add_numbered_paragraphs(doc, [data['address'], 'WYDZIAŁ FINANSÓW W MIEJSCU', 'a/a'], 'List Number 3', Inches(0.5))
            
        paragraph = doc.add_paragraph('\nSprawę prowadzi:   Beata Andrzejewska tel. 61 8410 568')
        
        doc.save(f"KT.5410.7.{data['kt']}.2024.docx")

    def __str__(self) -> str:
        return '\n'.join(f'{key} - {value}' for key, value in self.results.items()) + '\n' + '-' * 50


class ReadPDF:
    def __init__(self, path: str, reverse: bool = False) -> None:
        self.path: str = path
        self.reverse: bool = reverse
        self.files_paths: List[str] = [
            os.path.join(self.path, f)
            for f in os.listdir(self.path)
            if os.path.isfile(os.path.join(self.path, f)) and f.endswith('.pdf')
        ]

    def read_pdf(self) -> None: 
        handlers = []
        for file_path in self.files_paths[18:19]:
            handlers.append(PDFHandler(file_path))
        self.handlers: List[PDFHandler] = handlers

    def write_to_excel(self, data: List[PDFHandler], excel_file_path: str) -> None:
        """
        Write data extracted from OCR to an Excel file.
        """
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Extracted data (OCR)'

        for idx, handler in enumerate(data, start=1):
            kt = handler.results
            sheet[f'A{idx}'] = kt['kt']
            sheet[f'B{idx}'] = kt['tr']
            sheet[f'C{idx}'] = kt['vin']
            sheet[f'D{idx}'] = kt['name']
            sheet[f'E{idx}'] = ''
            sheet[f'F{idx}'] = ''
            sheet[f'G{idx}'] = kt['date']
            sheet[f'H{idx}'] = ''
            sheet[f'I{idx}'] = ''
            sheet[f'J{idx}'] = ''
            sheet[f'K{idx}'] = kt['czynność']
            sheet[f'L{idx}'] = kt['basis']
        workbook.save(excel_file_path)


def add_numbered_paragraphs(doc, items, style_name, left_indent=None, space_after=None):
    """
    Add numbered paragraphs to the document with specified style and formatting.
    """
    for item in items:
        paragraph = doc.add_paragraph(item, style=style_name)
        if left_indent is not None:
            paragraph.paragraph_format.left_indent = left_indent
        if space_after is not None:
            paragraph.paragraph_format.space_after = space_after


if __name__ == '__main__':
    # test = PDFHandler('./d.pdf')
    # print(test)
    # # print(test.text)
    # test.create_docx()

    a = ReadPDF('./skany')
    a.read_pdf()
    a.write_to_excel(a.handlers, 'output_ocr.xlsx')
    for pdf in a.handlers:
        pdf.create_docx()


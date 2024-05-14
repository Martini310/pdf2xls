import os
import re
import json
from datetime import date
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
import PyPDF2

dotenv.load_dotenv()

pytesseract.pytesseract.tesseract_cmd = os.environ.get('TESSERACT_CMD')

CUSTOM_CONFIG = r'--oem 3 --psm 6 -l pol'
poppler_path=os.environ.get('POPPLER_PATH')


class PDFHandler:
    """
    A class to read text from a PDF file, extract certain data patterns, and create a docx file.
    
    Args:
        path (str): The path to the PDF file.
        scan (bool, optional): Whether to perform OCR on scanned PDFs. 
                            Read text from PDF if false. Defaults to False.
    """
    name_ptrn = r'A-Za-zĄąĆćĘęŁłŃńÓóŚśŹźŻż”\"\'©—&-'
    patterns: Dict[str, Pattern[str]] = {
        'kt': r'(?<=KT.5410.[0-9].)\s*([0-9]+)',
        'name': rf'(?<=na rzecz )([{name_ptrn}]+(?:\s+[{name_ptrn}]+)*(?:\s+[{name_ptrn}]+))',
        'vin': r'(?<=VIN:)[\s —-]*([A-Z0-9—-]*)',
        'basis': r'(?<=w związku z art\. )[\w\s\.]*(?= ustawy)',
        'tr': r'(?<=rej\.)\s*([A-Z0-9]+\s*[A-Z0-9]+)\b',
        'address': r'(?<=na rzecz )[\s\w,.©\[\]/\\-]*(?=w związku)',
        'date': r'(?<=Poznań, dnia ).+(?=r)',
        'brand': r'(?<=marki\s)[\w\s\\/-—]+(?=o)',
        'pesel': r'[0-9]{9,11}',
        'purchase_date': r'(?<=z[\s]dnia)\s*[0-9-—/.\s]+(?=r.)',
    }

    def __init__(self, path: str, scan: bool = False) -> None:
        self.path: str = path
        self.scan: bool = scan
        if self.scan:
            self.text = self.perform_ocr(self.create_images(self.path))
        else:
            self.text = self.extract_text_from_pdf(self.path)

        self.results = self.extract_text(self.text, self.patterns)
        self.przypisz_czynnosc()
        self.kt_formatter()
        self.date_formatter(self.results.get('date'), 'date')
        self.date_formatter(self.results.get('purchase_date'), 'purchase_date')
        self.vin_formatter()

    def extract_text_from_pdf(self, file: str) -> str:
        """
        Extract text from PDF file (PDF with text, not a scanned file)

        Args:
            file (str): The path to the PDF file.

        Returns:
            str: Extracted text from the PDF.
        """
        pdf_text = ""
        with open(file, "rb") as f:
            pdf_reader = PyPDF2.PdfReader(f)
            for n, _ in enumerate(pdf_reader.pages):
                page = pdf_reader.pages[n]
                pdf_text += page.extract_text()
        return pdf_text

    def create_images(self, file: str) -> List[Image.Image]:
        """
        Convert PDF (scanned file) file into image.

        Args:
            file (str): The path to the PDF file.

        Returns:
            List[Image.Image]: List of PIL Image objects.
        """
        image: List[Image.Image] = pdf2image.convert_from_path(file, poppler_path=poppler_path)

        return image

    def perform_ocr(self, image: List[Image.Image]) -> str:
        """
        Perform OCR (Optical Character Recognition) on an image.

        Args:
            image (List[Image.Image]): List of PIL Image objects.

        Returns:
            str: Extracted text from the images.
        """
        extracted_text = []
        for page in image:
            text = pytesseract.image_to_string(page, config=CUSTOM_CONFIG)
            extracted_text.append(text)

        return '\n'.join(extracted_text)

    def find_pattern(self, text: str, pattern: Pattern[str]) -> str:
        """
        Find and return provided pattern in a given text.

        Args:
            text (str): The text to search the pattern in.
            pattern (Pattern[str]): Regular expression pattern.

        Returns:
            str: The found pattern.
        """
        try:
            matches: List[str] = re.findall(pattern, text)
            if not matches:
                return 'null'
            result = matches[0]
            result = result.strip().strip('.').strip(',').replace('\n', ' ')
            return result
        except (IndexError, TypeError):
            return 'błąd'

    def extract_text(self, text: str, patterns: Dict[str, Pattern[str]]) -> Dict[str, str]:
        """
        Return a dict with extracted data from given text based on patterns provided in dict.

        Args:
            text (str): The text to extract data from.
            patterns (Dict[str, Pattern[str]]): Dictionary of data patterns.

        Returns:
            Dict[str, str]: Extracted data.
        """
        data: Dict[str, str] = {}
        for key, pattern in patterns.items():
            extracted_text = self.find_pattern(text, pattern)
            data[key] = extracted_text
        return data

    def przypisz_czynnosc(self) -> None:
        '''
        Przypisuje czynność na podstawie podstawy prawnej przywołanej w postanowieniu.
        '''
        czynnosc = 'n/d'
        if '73aa ust. 1 pkt 3' in self.results['basis']:
            czynnosc = 'SPROWADZONY'
        elif '73aa ust. 1 pkt 1' in self.results['basis']:
            czynnosc = 'NABYCIE'
        elif '78 ust. 2 pkt 1' in self.results['basis']:
            czynnosc = 'ZBYCIE'
        self.results['czynność'] = czynnosc

    def kt_formatter(self) -> None:
        '''
        Fill kt number with zeros on the left if it's shorter than 5 and override it
        '''
        kt = self.results['kt']
        if kt != 'null' and len(kt) < 5:
            kt = kt.zfill(5)
        self.results['kt'] = kt

    def date_formatter(self, dt: str, name: str) -> None:
        '''
        Replace wrong characters in date and override it in self.results
        
        Args:
            dt (str): Date as a string
            name (str): Name of the date in self.results dict.
        '''
        if dt == 'null' or dt is None:
            return
        dt = dt.replace('—', '.').replace('-', '.').replace('/', '.')
        self.results[name] = dt

    def vin_formatter(self) -> None:
        '''
        Replace common mistakes in VIN pattern
        '''
        vin = self.results['vin']
        if vin == 'null':
            return
        vin = vin.replace('O', '0')
        self.results['vin'] = vin

    def create_docx(self) -> None:
        '''
        Create an administrative decision imposing a penalty in .docx format
        '''
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
            uzasadnienie += source['uzasadnienie_wspolne']
        elif self.results['czynność'] == 'SPROWADZONY':
            kara = source['kara_ue']
            uzasadnienie = source['uzasadnienie_ue']
            uzasadnienie += source['uzasadnienie_wspolne']
        elif self.results['czynność'] == 'ZBYCIE':
            kara = source['kara_zbycie']
            uzasadnienie = source['uzasadnienie_zbycie']
            uzasadnienie += source['uzasadnienie_wspolne'][3:]
            source['podstawa_prawna'] = source['podstawa_prawna'].replace('140mb ust. 1', '140mb ust. 6')

        today = date.today()
        formatted_date = f"Poznań dnia {today.strftime('%d.%m.%Y')}r."

        doc.add_paragraph(formatted_date).paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        header = 'Starosta Poznański\nul. Jackowskiego 18\n60-509 Poznań'
        doc.add_paragraph(header).paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

        doc.add_paragraph(f"\n\nDECYZJA NR KT.5410.7.{data['kt']}.2024\n", style=title)

        doc.add_paragraph(source['podstawa_prawna'].format(
            data['basis'], data['name'], data['pesel'], data['brand'], data['tr'], data['vin']
            ))

        doc.add_paragraph('\nStarosta\n', style=title)
        doc.add_paragraph(kara)
        doc.add_paragraph('\nUzasadnienie\n', style=title)

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

        doc.add_paragraph('\nPouczenie\n', style=title)
        doc.add_paragraph(source['pouczenia'][0])
        paragraph = doc.add_paragraph('\n')

        paragraph.add_run('\tWpłaty należy dokonać na konto numer: ')
        paragraph.add_run('7710 3012 4700 0000 0034 9162 41').bold = True
        paragraph.add_run(' w tytule podając nr decyzji ')
        paragraph.add_run(f"KT.5410.7.{data['kt']}.2024").bold = True
        paragraph = doc.add_paragraph()

        doc.add_paragraph(source['pouczenia'][2])
        doc.add_paragraph()
        doc.add_paragraph(source['pouczenia'][3])
        doc.add_paragraph('\n' * 9)

        doc.add_paragraph('Otrzymują:')

        receivers = [data['address'], 'WYDZIAŁ FINANSÓW W MIEJSCU', 'a/a']
        add_numbered_paragraphs(doc, receivers, 'List Number 3', Inches(0.5))

        doc.add_paragraph('\nSprawę prowadzi:   Beata Andrzejewska tel. 61 8410 568')

        if data['kt'] == 'null':
            file_name = f"KT.5410.7.{data['tr']}.2024.docx"
        else:
            file_name = f"KT.5410.7.{data['kt']}.2024.docx"

        doc.save(file_name)

    def __str__(self) -> str:
        return '\n'.join(f'{key} - {value}' for key,value in self.results.items()) + '\n' + '-' * 50


class ReadPDF:
    """
    A class to read multiple PDF files, extract data using PDFHandler, and write to Excel.

    Args:
        path (str): The path to the directory containing PDF files.
        scan (bool, optional): Whether to perform OCR on scanned PDFs. Defaults to False.
        reverse (bool, optional): Whether to reverse the order of extracted data. Defaults to False.
    """
    def __init__(self, path: str, scan: bool = False, reverse: bool = False) -> None:
        self.path: str = path
        self.reverse: bool = reverse
        self.scan: bool = scan
        self.handlers: List[PDFHandler] = None
        self.files_paths: List[str] = [
            os.path.join(self.path, f)
            for f in os.listdir(self.path)
            if os.path.isfile(os.path.join(self.path, f)) and f.endswith('.pdf')
        ]

    def read_pdf(self) -> None:
        '''
        Create a list of PDFHandler objects and store it in self.handlers
        '''
        handlers = []
        for file_path in self.files_paths:
            handlers.append(PDFHandler(file_path, scan=self.scan))
        self.handlers = handlers

    def write_to_excel(self, data: List[PDFHandler], excel_file_path: str) -> None:
        """
        Write data extracted from OCR to an Excel file.

        Args:
            data (List[PDFHandler]): List of PDFHandler objects containing extracted data.
            excel_file_path (str): The path to save the Excel file.
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
    # test = PDFHandler('./skany/test/2024-05-08-14-38-17-459_00008.pdf')
    # print(test)
    # print(test.text)
    # test.create_docx()

    a = ReadPDF('./skany/test2')
    a.read_pdf()
    a.write_to_excel(a.handlers, 'test.xlsx')
    for pdf in a.handlers:
        pdf.create_docx()

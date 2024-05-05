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

        doc.add_paragraph(f'Poznań dnia {date.today()}').paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        header = 'Starosta Poznański\nul. Jackowskiego 18\n60-509 Poznań'
        doc.add_paragraph(header).paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

        paragraph = doc.add_paragraph(f"\nDECYZJA NR KT.5410.7.{data['kt']}.2024\n", style=title)

        doc.add_paragraph(source['podstawa_prawna'].format(data['name'], data['pesel'], data['brand'], data['tr'], data['vin']))

        paragraph = doc.add_paragraph('\nStarosta\n', style=title)
        paragraph = doc.add_paragraph(source['kara'])
        paragraph = doc.add_paragraph('\nUzasadnienie\n', style=title)

        doc.add_paragraph(source['uzasadnienia'][0].format(data['name'], data['purchase_date']))
        doc.add_paragraph(source['uzasadnienia'][1])
        doc.add_paragraph(source['uzasadnienia'][2].format(data['date']))
        for uzasadnienie in source['uzasadnienia'][3:]:
            doc.add_paragraph(uzasadnienie)

        add_numbered_paragraphs(doc, source['przepisy'], 'ListNumber', Inches(0.5))

        for przepis in source['przepisy_2']:
            doc.add_paragraph(przepis)

        add_numbered_paragraphs(doc, source['przepisy_3'], 'List Number 2', space_after=Cm(0))
        paragraph = doc.add_paragraph(source['przepisy_4'])
        # add_numbered_paragraphs(doc, source['przepisy_5'], "List Number 3", Inches(0.5), Cm(0))
        for przepis in source['przepisy_5']:
            paragraph = doc.add_paragraph(przepis)
            paragraph.paragraph_format.left_indent = Inches(0.25)

        for przepis in source['przepisy_6']:
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
        paragraph = doc.add_paragraph('\n' * 11)
        
        paragraph = doc.add_paragraph('Otrzymują:')

        add_numbered_paragraphs(doc, [data['address'], 'WYDZIAŁ FINANSÓW W MIEJSCU', 'a/a'], 'List Number 3', Inches(0.5))
        # for idx, v in enumerate([data['address'], 'WYDZIAŁ FINANSÓW W MIEJSCU', 'a/a'], start=1):
        #     doc.add_paragraph(str(idx) + '.   ' + v)
            
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
        for file_path in self.files_paths[20:25]:
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

# Usage:
# write_to_excel_from_ocr(sentences, 'output_ocr.xlsx')


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

def create_docx():
    name = 'Martin Brzeziński'
    pesel = '293793759'
    marka = 'ford'
    tr = 'PZ 12345'
    vin = 'JFKKSAFHYJK98987'

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

    doc.add_paragraph('Poznań dnia ').paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    header = 'Starosta Poznański\nul. Jackowskiego 18\n60-509 Poznań'
    doc.add_paragraph(header).paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    paragraph = doc.add_paragraph('\nDECYZJA NR\n', style=title)

    podstawa_prawna = f"""Na podstawie art. 104 ustawy z dnia 14 czerwca 1960 r. – kodeks postepowania administracyjnego (Dz. U. z 2023 r., poz. 775 t. j.) w związku z art. 140mb ust. 1 oraz art. 73aa ust. 1 pkt 1 ustawy z dnia 20 czerwca 1997 r. – Prawo o ruchu drogowym (Dz. U. z 2023 r., poz. 1047 t. j.) po rozpatrzeniu sprawy Pana/i {name}. (PESEL: {pesel}) będącego właścicielem pojazdu marki {marka}. o numerze rejestracyjnym {tr}, nr nadwozia: {vin}."""

    doc.add_paragraph(podstawa_prawna)

    paragraph = doc.add_paragraph('\nStarosta\n', style=title)

    kara = """nakłada karę pieniężną w wysokości 500 zł (słownie: pięćset zł) w związku z niedopełnieniem obowiązku złożenia wniosku o rejestrację w terminie 30 dni od dnia nabycia wyżej wymienionego pojazdu, tj. obowiązku wynikającego z art. 73aa ust. 1 pkt 1 ustawy - Prawo o ruchu drogowym."""

    paragraph = doc.add_paragraph(kara)

    paragraph = doc.add_paragraph('\nUzasadnienie\n', style=title)

    uzasadnienia = ["\tTutejszy organ powziął informację z urzędu o tym, że strona w postępowaniu nie złożyła w terminie wniosku o rejestrację pojazdu nabytego na terytorium Rzeczpospolitej Polskiej. Z treści umowy/faktury nr ………….. pomiędzy ………………………. (sprzedającym) a …………………………. (kupującym) wynika, że strona nabyła pojazd w dniu ………………………….. r.",
        "\tZgodnie z art. 73aa ust. 1 pkt 1 ustawy Prawo o ruchu drogowym właściciel pojazdu jest obowiązany złożyć wniosek o jego rejestrację w terminie nieprzekraczającym 30 dni od dnia jego nabycia na terytorium Rzeczpospolitej Polskiej.",
        "\tW związku z niedopełnieniem wyrażonego w ustawie – Prawo o ruchu drogowym obowiązku, tutejszy organ wszczął z urzędu w dniu ………………………. r. postępowanie administracyjne w przedmiocie wyżej wskazanego naruszenia o czym pisemnie zawiadomił stronę. Skutecznie doręczone zawiadomienie o wszczęciu postepowania umożliwiło stronie czynny udział w toczącym się postępowaniu i wypowiedzenie się w przedmiotowej sprawie. Strona nie złożyła pisemnego wyjaśnienia.",
        "\tZgodnie z art. 140mb ust. 1 ustawy Prawo o ruchu drogowym, kto będąc właścicielem pojazdu obowiązanym do złożenia wniosku o rejestracje pojazdu w terminie, o którym mowa w art. 73aa ust. 1, nie złoży tego wniosku w terminie, podlega karze w wysokości 500 zł.",
        "\tMając na uwadze powyższe organ ustalił, że zasadne jest zastosowanie kary w wysokości 500 zł (słownie: pięćset zł).",
        "\tW myśl art. 140n ust. 6 Prawo o ruchu drogowym do kar pieniężnych, o których mowa w art. 140ma i art. 140mb, nie stosuje się  przepisów art. 189d-189f ustawy z dnia 14 czerwca 1960 r. –Kodeks postepowania administracyjnego tj.:",
        "art. 189d wymierzając administracyjną karę pieniężną, organ administracji publicznej bierze pod uwagę:",
    ]

    przepisy = ["wagę i okoliczności naruszenia prawa, w szczególności potrzebę ochrony życia lub zdrowia, ochrony mienia w znacznych rozmiarach lub ochrony ważnego interesu publicznego lub wyjątkowo ważnego interesu strony oraz czas trwania tego naruszenia;",
        "częstotliwość niedopełniania w przeszłości obowiązku albo naruszania zakazu tego samego rodzaju co niedopełnienie obowiązku albo naruszenie zakazu, w następstwie którego ma być nałożona kara;",
        "uprzednie ukaranie za to samo zachowanie za przestępstwo, przestępstwo skarbowe, wykroczenie lub wykroczenie skarbowe;",
        "stopień przyczynienia się strony, na którą jest nakładana administracyjna kara pieniężna, do powstania naruszenia prawa;",
        "działania podjęte przez stronę dobrowolnie w celu uniknięcia skutków naruszenia prawa;",
        "wysokość korzyści, którą strona osiągnęła, lub straty, której uniknęła;",
        "w przypadku osoby fizycznej - warunki osobiste strony, na którą administracyjna kara pieniężna jest nakładana;",
    ]

    przepisy_2 = ["art. 189e w przypadku gdy do naruszenia prawa doszło wskutek działania siły wyższej, strona nie podlega ukaraniu;",
        "art. 189f",
        "1. organ administracji publicznej, w drodze decyzji, odstępuje od nałożenia administracyjnej kary pieniężnej i poprzestaje na pouczeniu, jeżeli:"
        ]

    przepisy_3 = ["waga naruszenia prawa jest znikoma, a strona zaprzestała naruszania prawa lub",
        "za to samo zachowanie prawomocną decyzją na stronę została uprzednio nałożona administracyjna kara pieniężna przez inny uprawniony organ administracji publicznej lub strona została prawomocnie ukarana za wykroczenie lub wykroczenie skarbowe, lub prawomocnie skazana za przestępstwo lub przestępstwo skarbowe i uprzednia kara spełnia cele, dla których miałaby być nałożona administracyjna kara pieniężna.",
    ]

    przepisy_4 = """2. w przypadkach innych niż wymienione w § 1, jeżeli pozwoli to na spełnienie celów, dla których miałaby być nałożona administracyjna kara pieniężna, organ administracji publicznej, w drodze postanowienia, może wyznaczyć stronie termin do przedstawienia dowodów potwierdzających:"""

    przepisy_5 = ["usunięcie naruszenia prawa lub",
        "powiadomienie właściwych podmiotów o stwierdzonym naruszeniu prawa, określając termin i sposób powiadomienia."
        ]

    przepisy_6 = ["3. organ administracji publicznej w przypadkach, o których mowa w § 2, odstępuje od nałożenia administracyjnej kary pieniężnej i poprzestaje na pouczeniu, jeżeli strona przedstawiła dowody, potwierdzające wykonanie postanowienia.",
        "W związku z powyższym przywołany przepis art. 140n ust. 6 wyklucza możliwość odstąpienia od nałożenia kary pieniężnej i obniżenia jej wysokości.",
        "W tej sytuacji orzeka się jak w sentencji."
    ]

    pouczenia = [
        "\tZgodnie z art. 140n ust 5 ustawy Prawo o ruchu drogowym kary pieniężne są wnoszone na rachunek bankowy starostwa w terminie 14 dni od dnia, w którym decyzja o nałożeniu kary pieniężnej stała się ostateczna.",
        "\tWpłaty należy dokonać na konto numer: 7710 3012 4700 0000 0034 9162 41 w tytule podając nr decyzji KT.5410.7.00049.2024",
        "\tOd niniejszej decyzji przysługuje odwołanie do Samorządowego Kolegium Odwoławczego w Poznaniu, za pośrednictwem Starosty Poznańskiego, w terminie 14 dni od daty jej doręczenia.",
        "\tW trakcie biegu terminu od wniesienia odwołania stronie służy także prawo do zrzeczenia się prawa do wniesienia odwołania od decyzji. Z dniem doręczenia organowi oświadczenia o zrzeczeniu się prawa do wniesienia odwołania, decyzja staje się ostateczna i prawomocna.",
    ]

    for uzasadnienie in uzasadnienia:
        doc.add_paragraph(uzasadnienie)

    add_numbered_paragraphs(doc, przepisy, 'ListNumber', Inches(0.5))

    for przepis in przepisy_2:
        doc.add_paragraph(przepis)

    add_numbered_paragraphs(doc, przepisy_3, 'List Number 2', space_after=Cm(0))

    paragraph = doc.add_paragraph(przepisy_4)

    add_numbered_paragraphs(doc, przepisy_5, "List Number 3", Inches(0.5), Cm(0))

    for przepis in przepisy_6:
        paragraph = doc.add_paragraph(przepis)

    paragraph = doc.add_paragraph('\nPouczenie\n', style=title)

    paragraph = doc.add_paragraph(pouczenia[0])
    paragraph = doc.add_paragraph('\n')

    paragraph.add_run('\tWpłaty należy dokonać na konto numer: ')
    paragraph.add_run('7710 3012 4700 0000 0034 9162 41').bold = True
    paragraph.add_run(' w tytule podając nr decyzji ')
    paragraph.add_run('KT.5410.7.00049.2024').bold = True
    paragraph = doc.add_paragraph()
    
    paragraph = doc.add_paragraph(pouczenia[2])
    paragraph = doc.add_paragraph()
    paragraph = doc.add_paragraph(pouczenia[3])
    paragraph = doc.add_paragraph('\n' * 11)
    
    paragraph = doc.add_paragraph('Otrzymują:')

    # add_numbered_paragraphs(doc, [name, 'WYDZIAŁ FINANSÓW W MIEJSCU', 'a/a'], 'MyNumberedList', Inches(0.5))
    for idx, v in enumerate([name, 'WYDZIAŁ FINANSÓW W MIEJSCU', 'a/a'], start=1):
        doc.add_paragraph(str(idx) + '.   ' + v)
        
    paragraph = doc.add_paragraph('\nSprawę prowadzi:   Beata Andrzejewska tel. 61 8410 568')
    
    doc.save('test.docx')


if __name__ == '__main__':
    test = PDFHandler('./d.pdf')
    print(test)
    # print(test.text)
    test.create_docx()

    # a = ReadPDF('./skany')
    # a.read_pdf()
    # a.write_to_excel(a.handlers, 'output_ocr.xlsx')
    # for pdf in a.handlers:
    #     pdf.create_docx()


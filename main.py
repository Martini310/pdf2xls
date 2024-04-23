import os
import re
from typing import List, Dict
from PIL import Image
import pdf2image
import dotenv
from openpyxl import Workbook
from docx import Document
from docx.shared import Pt, Cm
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

    section = doc.sections[0]

    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    paragraph = doc.add_paragraph('Poznań dnia ')

    paragraph.paragraph_format.line_spacing = 1
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT


    paragraph = doc.add_paragraph('Starosta Poznański\nul. Jackowskiego 18\n60-509 Poznań')

    paragraph.paragraph_format.line_spacing = 1
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT


    paragraph = doc.add_paragraph('\nDECYZJA NR ')

    paragraph.paragraph_format.line_spacing = 1
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.runs[0].bold = True


    name = 'Martin Brzeziński'
    pesel = '293793759'
    marka = 'ford'
    tr = 'PZ 12345'
    vin = 'JFKKSAFHYJK98987'

    podstawa_prawna = f"""Na podstawie art. 104 ustawy z dnia 14 czerwca 1960 r. – kodeks postepowania administracyjnego (Dz. U. z 2023 r., poz. 775 t. j.) w związku z art. 140mb ust. 1 oraz art. 73aa ust. 1 pkt 1 ustawy z dnia 20 czerwca 1997 r. – Prawo o ruchu drogowym (Dz. U. z 2023 r., poz. 1047 t. j.) po rozpatrzeniu sprawy Pana/i {name}. (PESEL: {pesel}) będącego właścicielem pojazdu marki {marka}. o numerze rejestracyjnym {tr}, nr nadwozia: {vin}."""

    paragraph = doc.add_paragraph(podstawa_prawna)
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


    paragraph = doc.add_paragraph('Starosta')
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.runs[0].bold = True

    kara = """nakłada karę pieniężną w wysokości 500 zł (słownie: pięćset zł) w związku z niedopełnieniem obowiązku złożenia wniosku o rejestrację w terminie 30 dni od dnia nabycia wyżej wymienionego pojazdu, tj. obowiązku wynikającego z art. 73aa ust. 1 pkt 1 ustawy - Prawo o ruchu drogowym."""

    paragraph = doc.add_paragraph(kara)
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    paragraph = doc.add_paragraph('Uzasadnienie')
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.runs[0].bold = True
    
    uzasadnienia = ["Tutejszy organ powziął informację z urzędu o tym, że strona w postępowaniu nie złożyła w terminie wniosku o rejestrację pojazdu nabytego na terytorium Rzeczpospolitej Polskiej. Z treści umowy/faktury nr ………….. pomiędzy ………………………. (sprzedającym) a …………………………. (kupującym) wynika, że strona nabyła pojazd w dniu ………………………….. r.",
    "  Zgodnie z art. 73aa ust. 1 pkt 1 ustawy Prawo o ruchu drogowym właściciel pojazdu jest obowiązany złożyć wniosek o jego rejestrację w terminie nieprzekraczającym 30 dni od dnia jego nabycia na terytorium Rzeczpospolitej Polskiej.",
    "  W związku z niedopełnieniem wyrażonego w ustawie – Prawo o ruchu drogowym obowiązku, tutejszy organ wszczął z urzędu w dniu ………………………. r. postępowanie administracyjne w przedmiocie wyżej wskazanego naruszenia o czym pisemnie zawiadomił stronę. Skutecznie doręczone zawiadomienie o wszczęciu postepowania umożliwiło stronie czynny udział w toczącym się postępowaniu i wypowiedzenie się w przedmiotowej sprawie. Strona nie złożyła pisemnego wyjaśnienia.",
    "  Zgodnie z art. 140mb ust. 1 ustawy Prawo o ruchu drogowym, kto będąc właścicielem pojazdu obowiązanym do złożenia wniosku o rejestracje pojazdu w terminie, o którym mowa w art. 73aa ust. 1, nie złoży tego wniosku w terminie, podlega karze w wysokości 500 zł.",
    "  Mając na uwadze powyższe organ ustalił, że zasadne jest zastosowanie kary w wysokości 500 zł (słownie: pięćset zł).",
    "  W myśl art. 140n ust. 6 Prawo o ruchu drogowym do kar pieniężnych, o których mowa w art. 140ma i art. 140mb, nie stosuje się  przepisów art. 189d-189f ustawy z dnia 14 czerwca 1960 r. –Kodeks postepowania administracyjnego tj.:",
    "  art. 189d wymierzając administracyjną karę pieniężną, organ administracji publicznej bierze pod uwagę:"]

    
    for uzasadnienie in uzasadnienia:
        par1 = doc.add_paragraph(uzasadnienie)
        par1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # paragraph.paragraph_format.line_spacing = 1
        par1.paragraph_format.space_before = Cm(0)
        par1.paragraph_format.space_before = Cm(0)
    
    przepisy = ["wagę i okoliczności naruszenia prawa, w szczególności potrzebę ochrony życia lub zdrowia, ochrony mienia w znacznych rozmiarach lub ochrony ważnego interesu publicznego lub wyjątkowo ważnego interesu strony oraz czas trwania tego naruszenia;",
        "częstotliwość niedopełniania w przeszłości obowiązku albo naruszania zakazu tego samego rodzaju co niedopełnienie obowiązku albo naruszenie zakazu, w następstwie którego ma być nałożona kara;",
        "uprzednie ukaranie za to samo zachowanie za przestępstwo, przestępstwo skarbowe, wykroczenie lub wykroczenie skarbowe;",
        "stopień przyczynienia się strony, na którą jest nakładana administracyjna kara pieniężna, do powstania naruszenia prawa;",
        "działania podjęte przez stronę dobrowolnie w celu uniknięcia skutków naruszenia prawa;",
        "wysokość korzyści, którą strona osiągnęła, lub straty, której uniknęła;",
        "w przypadku osoby fizycznej - warunki osobiste strony, na którą administracyjna kara pieniężna jest nakładana;",
    ]
    
    for item in przepisy:
        paragraph = doc.add_paragraph(item)

        # Set the numbering style to '1, 2, 3' (ordered list)
        paragraph.style = 'ListNumber'
        
        
    przepisy_2 = """art. 189e w przypadku gdy do naruszenia prawa doszło wskutek działania siły wyższej, strona nie podlega ukaraniu;
art. 189f
1. organ administracji publicznej, w drodze decyzji, odstępuje od nałożenia administracyjnej kary pieniężnej i poprzestaje na pouczeniu, jeżeli:
"""
    paragraph = doc.add_paragraph(przepisy_2)
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY
        
    przepisy_3 = ["waga naruszenia prawa jest znikoma, a strona zaprzestała naruszania prawa lub",
        "za to samo zachowanie prawomocną decyzją na stronę została uprzednio nałożona administracyjna kara pieniężna przez inny uprawniony organ administracji publicznej lub strona została prawomocnie ukarana za wykroczenie lub wykroczenie skarbowe, lub prawomocnie skazana za przestępstwo lub przestępstwo skarbowe i uprzednia kara spełnia cele, dla których miałaby być nałożona administracyjna kara pieniężna.",
    ]
    
    for item in przepisy_3:
        paragraph = doc.add_paragraph(item)
        paragraph.style = 'ListNumber'
        
    przepisy_4 = """2. w przypadkach innych niż wymienione w § 1, jeżeli pozwoli to na spełnienie celów, dla których miałaby być nałożona administracyjna kara pieniężna, organ administracji publicznej, w drodze postanowienia, może wyznaczyć stronie termin do przedstawienia dowodów potwierdzających:"""
    
    paragraph = doc.add_paragraph(przepisy_4)
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY
    
    przepisy_5 = ["usunięcie naruszenia prawa lub",
        "powiadomienie właściwych podmiotów o stwierdzonym naruszeniu prawa, określając termin i sposób powiadomienia."
        ]
    for item in przepisy_5:
        paragraph = doc.add_paragraph(item)
        paragraph.style = 'ListNumber'
    
    przepisy_6 = """3. organ administracji publicznej w przypadkach, o których mowa w § 2, odstępuje od nałożenia administracyjnej kary pieniężnej i poprzestaje na pouczeniu, jeżeli strona przedstawiła dowody, potwierdzające wykonanie postanowienia.
	W związku z powyższym przywołany przepis art. 140n ust. 6 wyklucza możliwość odstąpienia od nałożenia kary pieniężnej i obniżenia jej wysokości.
	W tej sytuacji orzeka się jak w sentencji.
    """
    paragraph = doc.add_paragraph(przepisy_6)
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY
    
    paragraph = doc.add_paragraph('\nPouczenie\n')
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.runs[0].bold = True
    
    
    
    
    
    doc.save('test.docx')
    
create_docx()

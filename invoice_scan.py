# to scan invoices
# this was written by GPT so there's going to be some trouble shooting involved

import cv2
import pypdf
import pytesseract
import pdf2image
import os, re, datetime
import openpyxl
from pdf2image import convert_from_path
from openpyxl import Workbook, load_workbook
from datetime import date

# You must install tesseract on your machine and update the path below
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Function to perform OCR
def ocr_core(image):
    """
    This function will handle the OCR processing of images.
    """
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Apply OCR to the preprocessed image
    text = pytesseract.image_to_string(gray)

    return text

def extract_text_from_pdf(file_path):
    pdf_file_obj = open(file_path, 'rb')
    pdf_reader = pypdf.PdfReader(pdf_file_obj)
    num_pages = len(pdf_reader.pages)
    text = ""
    for page in range(num_pages):
        #page_obj = pdf_reader.getPage(page)
        page_obj = pdf_reader.pages[page]
        text += page_obj.extract_text()
    pdf_file_obj.close()
    return text

def is_file_in_folder(filename, folder_path):
    file_path = os.path.join(folder_path, filename)
    return os.path.isfile(file_path)



file_name = f'{date.today()}_invoices.xlsx'

folder_path = r'C:\Users\Michael\Desktop\python-work\Invoices'

if is_file_in_folder(file_name, folder_path):
    wb = load_workbook(os.path.join(folder_path, file_name))
    ws = wb.active
else:
    wb = Workbook()

    ws = wb.active

    ws['A1'] = 'Vendor'
    ws['B1'] = 'Date'
    ws['C1'] = 'Number'
    ws['D1'] = 'Amount Due'

invoice_counter = len(ws['C'])
unique_invoices = []

for cell in ws['C']:
    unique_invoices.append(cell.value)

## Loops through the files in a chosen folder, extracts the relevant text from them
## and puts them in an excel file
for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"): # add any file type you want to process
        print(os.path.join(folder_path, filename))
        # Replace 'path_to_your_pdf' with the path to the PDF file you want to process
        """ text = extract_text_from_pdf(os.path.join(folder_path, filename))
        print(text) """
        
        pages = convert_from_path(os.path.join(folder_path, filename), dpi=300)

        # Perform OCR on each image
        invoice_no = None

        for i, page in enumerate(pages):
            match_counter = 0
            page.save('out.jpg', 'JPEG')
            img = cv2.imread('out.jpg')
            print(f"Content on {ocr_core(img)}")
            content = ocr_core(img)
            
            ## Finds the invoice number and checks for uniqueness
            inv_pattern = r'INVOICE\n\n(\d+-\d+)'
            inv_pattern2 = r'office (\d+-?\d*)'
            page_pattern = r"page\s(\d+)\sof\s([2-9])"



            inv_match = re.search(inv_pattern, content, re.IGNORECASE)

            if inv_match:
                match_counter = 1
            inv_match2 = re.search(inv_pattern2, content, re.IGNORECASE)
            if inv_match2:
                match_counter = 2
            page_match = re.search(page_pattern, content, re.IGNORECASE)
            print(inv_match)
            if match_counter:
                if match_counter == 1:
                    invoice_no = inv_match.group(1)
                else:
                    invoice_no = inv_match2.group(1)

                print(f"Found 'invoice number': {invoice_no}")
                if invoice_no not in unique_invoices:
                    invoice_counter += 1
                    unique_invoices.append(invoice_no)
                    ws[f'C{invoice_counter}'] = f'{invoice_no}'
                    
                    # Finds the vendor
                    if content.startswith("GEN") or content.startswith("GEM"):
                        print("Vendor: GEMCO")
                        ws[f'A{invoice_counter}'] = 'Gemco'

                    # Finds the date of the invoice    
                    date_match = re.search(r'invoice date[:\s]*([01]?\d/[0123]?\d/\d{2})', content, re.IGNORECASE)
                    if date_match:
                        # If a match was found, 'group(1)' contains the first parenthesized subgroup - the date.
                        invoice_date = date_match.group(1)
                        print(f"Found 'invoice date': {invoice_date}")
                        ws[f'B{invoice_counter}'] = f'{invoice_date}'
                    else:
                        print("No 'invoice date' found")

                    ## Finds the invoice amount due
                    total_match = re.search(r'balance \$([\d,]+\.\d{2})', content, re.IGNORECASE)
                    if total_match:
                        # If a match was found, 'group(1)' contains the first parenthesized subgroup - the balance amount.
                        balance_amount = total_match.group(1)
                        print(f"Found 'balance amount': ${balance_amount}")
                        ws[f'D{invoice_counter}'].value = balance_amount
                    else:
                        print("No 'balance amount' found")
            if page_match:
                total_match = re.search(r'balance \$([\d,]+\.\d{2})', content, re.IGNORECASE)
                if total_match:
                    # If a match was found, 'group(1)' contains the first parenthesized subgroup - the balance amount.
                    balance_amount = total_match.group(1)
                    print(f"Found 'balance amount': ${balance_amount}")
                    ws[f'D{invoice_counter}'] = balance_amount
            else:
                print("No 'unique invoice number' found")
            
            

    else:
        continue

wb.save(os.path.join(folder_path, file_name))

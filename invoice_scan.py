# to scan invoices
# this was written by GPT so there's going to be some trouble shooting involved

import cv2
import pypdf
import pytesseract
import pdf2image
import os, re
from pdf2image import convert_from_path

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


folder_path = r'C:\Users\Michael\Desktop\python-work\Invoices'


for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"): # add any file type you want to process
        print(os.path.join(folder_path, filename))
        # Replace 'path_to_your_pdf' with the path to the PDF file you want to process
        """ text = extract_text_from_pdf(os.path.join(folder_path, filename))
        print(text) """
        pages = convert_from_path(os.path.join(folder_path, filename), dpi=300)

        # Perform OCR on each image
        for i, page in enumerate(pages):
            page.save('out.jpg', 'JPEG')
            img = cv2.imread('out.jpg')
            print(f"Content on page {i+1}:\n{ocr_core(img)}")
            content = ocr_core(img)
            # Finds the vendor
            if content.startswith("GEN"):
                print("Vendor: GEMCO")

            # Finds the date of the invoice    
            date_match = re.search(r'invoice date[:\s]*([01]?\d/[0123]?\d/\d{2})', content, re.IGNORECASE)
            if date_match:
                # If a match was found, 'group(1)' contains the first parenthesized subgroup - the date.
                invoice_date = date_match.group(1)
                print(f"Found 'invoice date': {invoice_date}")
            else:
                print("No 'invoice date' found")
            
            ## Finds the invoice number

            inv_match = re.search(r'main office (\d+-?\d*)', content, re.IGNORECASE)
            if inv_match:
                # If a match was found, 'group(1)' contains the first parenthesized subgroup - the office number.
                invoice_no = inv_match.group(1)
                print(f"Found 'main office number': {invoice_no}")
            else:
                print("No 'main office number' found")

            ## Finds the invoice amount due
            total_match = re.search(r'balance \$([\d,]+\.\d{2})', content, re.IGNORECASE)
            if total_match:
                # If a match was found, 'group(1)' contains the first parenthesized subgroup - the balance amount.
                balance_amount = total_match.group(1)
                print(f"Found 'balance amount': ${balance_amount}")
            else:
                print("No 'balance amount' found")

    else:
        continue

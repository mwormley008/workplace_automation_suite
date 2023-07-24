# to scan invoices
# this was written by GPT so there's going to be some trouble shooting involved

import cv2
import pytesseract
import pdf2image
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

# Convert PDF to images
pages = convert_from_path(r'C:\Users\Michael\Downloads\140338.pdf', dpi=300)

# Perform OCR on each image
for i, page in enumerate(pages):
    page.save('out.jpg', 'JPEG')
    img = cv2.imread('out.jpg')
    print(f"Content on page {i+1}:\n{ocr_core(img)}")

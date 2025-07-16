# to scan invoices
# this was written by GPT so there's going to be some trouble shooting involved


## TODO: I'm not finding the invoice information / invoice number on sheet metal supply or pfs invoices, looks like there is 
# legible invoice info on Stevenson

import cv2
import pypdf
import pytesseract
import pdf2image
import os, re, datetime
import openpyxl
import numpy as np
from pdf2image import convert_from_path
from openpyxl import Workbook, load_workbook
from datetime import date
from qbdetector import dual_print

from tkinter import Tk, simpledialog, messagebox

# You must install tesseract on your machine and update the path below
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Function to perform OCR
def ocr_core(image):
    """
    This function will handle the OCR processing of images.
    """
    gray = image

    # Get the height and width of the image
    height, width = gray.shape
    
    # Define the bottom right quadrant
    x = 2 * width // 4
    y = 3 * height // 4
    w = width // 2
    h = height // 4

    # Crop the region of interest for the invoice number
    inv_roi = gray[0:height//2, width//2:width]
    inv_text = pytesseract.image_to_string(inv_roi, config='--psm 6')

    # Crop the region of interest for the amount
    amt_roi = gray[y:y+h, x:x+w]
    
    amt_text = pytesseract.image_to_string(amt_roi, config='--psm 6')

    text = pytesseract.image_to_string(gray)

    return text, inv_text, amt_text

def preprocess_image(img):
    # Convert to grayscale (you're already doing this)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Thresholding
    _, thresh = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Noise removal
    kernel = np.ones((1, 1), np.uint8)
    opening = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)
    closing = cv2.morphologyEx(opening, cv2.MORPH_CLOSE, kernel)

    return closing


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


folder_path = r'C:\Users\Michael\Documents\ShineDoc\sources'
output_file = open("output.txt", "w")

for filename in os.listdir(folder_path):
    full_path = os.path.join(folder_path, filename)
    print(full_path)
    
    img = cv2.imread(full_path)

    # Check if the image was read successfully
    if img is None:
        print(f"Error reading the image at path: {full_path}")
        continue

    # Check if the image is grayscale
    if len(img.shape) == 2 or img.shape[2] == 1:
        # The image is already grayscale
        processed_img = img
    else:
        # Convert the image to grayscale
        processed_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Now, you can apply the preprocess_image function and other processing steps.



    content, inv_text, amt_text = ocr_core(processed_img)
    dual_print(content)
    dual_print(inv_text)
    dual_print(amt_text)
print("Done processing, check the log file for the outputs.")

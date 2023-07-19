# Script to put in the expense categories

import pyautogui, openpyxl, datetime, calendar, os
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta, date

from tkinter import Tk, simpledialog
from tkinter.filedialog import askopenfilename


# Create the Tkinter root window
root = Tk()
root.withdraw()  # Hide the root window
initial_dir=r"\\WBR\data\shared\G702 & G703 Forms"
workbook_path = askopenfilename(initialdir=initial_dir)
print(workbook_path)
"\\WBR\data\shared\G702 & G703 Forms\41 North Contractors Raising Canes New Lenox 15682 3.xlsx"
# Open AIA form based on invoice number
# maybe I could do this with tkinter idk
## If I have tkinter form I guess it could prompt old inv number,
## new invoice number, and total completed and stored to date
"""from tkinter import Tk
from tkinter.filedialog import askopenfilename


# Open a file dialog for selecting a file

# Print the selected file path
print("Selected file:", file_path)
When you run this code, a file dialog window will appear, allowing you to browse and select the file you want. Once you select the file and click "Open," the selected file path will be stored in the file_path variable, and you can use it in your Python code as needed.

Note that the askopenfilename() function returns the selected file path as a string. You can modify the code after obtaining the file path to perform further operations on the selected file.

Make sure to import the necessary modules (Tk and askopenfilename) from tkinter and create a Tk root window before using the askopenfilename() function.
"""
new_invoice_number = simpledialog.askinteger("Invoice Prompt", "Enter the new invoice number:")

# Page 1
workbook = load_workbook(filename=workbook_path)
sheet1 = workbook["G702"]

# Increment application number
application = sheet1["J3"]
old_application_number = application.value
print(application.value)
application.value = application.value + 1
new_application_number = application.value

# increment work period

today = date.today()
print(today)
res = calendar.monthrange(today.year, today.month)[1]
workperiod = sheet1["J7"]
workperiod.value = f"{today.month}/{res}/{today.year}"
print(workperiod)

# change invoice number to the new invoice number
# J9
invoice_number = sheet1["J9"]
old_invoice_number = invoice_number.value
invoice_number.value = new_invoice_number

# add old current payment E39 to previous payments E38
sheet1["E38"].value += sheet1["E39"].value
sheet1["E39"].value = 0

# Page 2
sheet2 = workbook["G703"]
prev_apps = ["D13", "D17", "D21", "D25"]
this_period = ["E13", "E17", "E21", "E25"]

for x, y in zip(prev_apps, this_period):
    sheet2[x].value += sheet2[y].value
    sheet2[y].value = 0


# Add old this period to previous applications
#D13 E13
# Each new row increases by 4 (D17 E17)

# Save a new file with new name
"""
new_workbook_path = workbook_path.replace(str(old_invoice_number), str(new_invoice_number))
new_workbook_path = workbook_path.replace(str(old_application_number), str(new_application_number))

print(new_workbook_path)
print(type(new_workbook_path))

new_workbook_path = os.path.normpath(new_workbook_path)
print(new_workbook_path)
workbook.save(new_workbook_path)"""

# Add formulae to all cells 
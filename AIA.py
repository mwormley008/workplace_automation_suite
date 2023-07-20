# Script to put in the expense categories

import pyautogui, openpyxl, datetime, calendar, os, sys, re
import subprocess
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta, date

from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename


# Create the Tkinter root window
root = Tk()
root.withdraw()  # Hide the root window

response = messagebox.askyesno("Confirmation", "Do you want to start in Quickbooks?")

if response:
    qtoken = 1
    exec(open('progress_invoice.py').read())
else:
    qtoken = 0


old_invoice_number = simpledialog.askinteger("Invoice Prompt", "Enter the old invoice number:")

initial_dir=r"\\WBR\data\shared\G702 & G703 Forms"

def find_file_by_number(folder_path, target_number):
    pattern = re.compile(r'.*{}.*\.xlsx$'.format(target_number))

    for filename in os.listdir(folder_path):
        if pattern.match(filename):
            file_path = os.path.join(folder_path, filename)
            return file_path

    # If the file is not found
    return None

# Example usage
folder_path = initial_dir # Replace with the actual folder path
target_number = old_invoice_number  # Replace with the specific number you are looking for

file_path = find_file_by_number(folder_path, target_number)

if file_path:
    print(f"File found: {file_path}")
    use_file = messagebox.askyesno("Confirmation", f"Do you want to use this file?/\n {file_path}")
    if use_file:
        workbook_path = file_path
    else:
        workbook_path = askopenfilename(initialdir=initial_dir)
else:
    print("File not found.")
    workbook_path = askopenfilename(initialdir=initial_dir)
print(workbook_path)


if qtoken == 0:
    new_inv = simpledialog.askinteger("Invoice Prompt", "Enter the new invoice number:")
    completed_through= simpledialog.askinteger("Invoice Prompt", "Enter the amount billed without retention taken out:")

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
invoice_number.value = new_inv
# add work from this period to completed and stored to date
completed_through_cell = sheet1["E26"]
completed_through_cell.value += completed_through

# add old current payment E39 to previous payments E38
retention_percentage = sheet1['B29'].value
sheet1["E38"].value += sheet1["E39"].value
sheet1["E39"].value = completed_through * (1-retention_percentage*.01)


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
new_workbook_path = workbook_path.replace(str(old_invoice_number), str(new_inv))
new_workbook_path = new_workbook_path.split(' ')
new_workbook_path[-1] = new_workbook_path[-1].replace(str(old_application_number), str(new_application_number))
new_workbook_path = (' ').join(new_workbook_path)
print(new_workbook_path)
#new_workbook_path = new_workbook_path.replace(str(old_application_number), str(new_application_number))


#print(new_workbook_path)
#print(type(new_workbook_path))

#print(new_workbook_path)
#workbook.save(new_workbook_path)
workbook.save(new_workbook_path)

# TODO: Add formulae to all cells

excel_path = r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"

subprocess.run([excel_path, new_workbook_path])

"""os.startfile(excel_path)
os.startfile(new_workbook_path)


print("start", '', '"' + new_workbook_path + '"')
subprocess.run(["start", '', '"' + new_workbook_path + '"'], shell=True)
"""

"""
child_process = subprocess.Popen(new_workbook_path)
child_process.wait()
"""
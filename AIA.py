# Script to put in the expense categories

import pyautogui, openpyxl, datetime, calendar, os, sys, re
import subprocess, time
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta, date

from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename, Frame, Button
from time import sleep

class CustomDialog(simpledialog.Dialog):
    def __init__(self, parent, title=None, button_names=None):
        self.button_names = button_names
        super().__init__(parent, title=title)

    def body(self, master):
        return None  # override body function to prevent an extra Entry widget

    def buttonbox(self):
        # add custom buttons
        box = Frame(self)
        for button in self.button_names:
            Button(box, text=button, height=5, width=15, command=lambda text=button: self.ok(text)).pack(side="left", padx=5, pady=5)
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)
        box.pack()

    def ok(self, button_name, event=None):
        self.result = button_name
        self.destroy()

def find_file_by_number(folder_path, target_number):
    pattern = re.compile(r'.*{}.*\.xlsx$'.format(target_number))

    for filename in os.listdir(folder_path):
        if pattern.match(filename):
            file_path = os.path.join(folder_path, filename)
            return file_path

    # If the file is not found
    return None

# Create the Tkinter root window
root = Tk()
root.withdraw()  # Hide the root window
noe_token = 0
new_or_existing = CustomDialog(root, title="New or Existing", button_names=["New Billing Cycle", "Existing Billing Cycle"])
print("You clicked:", new_or_existing.result)

if new_or_existing.result == 'New Billing Cycle':
    noe_token = 1

qb_response = messagebox.askyesno("Confirmation", "Do you want to start in Quickbooks?")
sleep(1)

# Going to need new_contract_invoice.py
if qb_response and noe_token == 0:
    qtoken = 1
    exec(open('progress_invoice.py').read())
elif qb_response and noe_token == 1:
    qtoken = 1
    contract_date = simpledialog.askinteger("Contract Prompt", "What is the contract date DD/MM/YY:")
    exec(open('new_contract_invoice.py').read())
elif qb_response == 0 and noe_token == 1:
    qtoken = 0
    contract_date = simpledialog.askstring("Contract Prompt", "What is the contract date DD/MM/YY:")
    bill_to = simpledialog.askstring("Contractor", "Who is the contractor?")
    ship_to = simpledialog.askstring("Project", "What is the project name?")

else:
    qtoken = 0


initial_dir=r"\\WBR\data\shared\G702 & G703 Forms"

folder_path = initial_dir 


if noe_token == 0:
    old_invoice_number = simpledialog.askinteger("Invoice Prompt", "Enter the old invoice number:")
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
else:
    workbook_path = r"\\WBR\data\shared\G702 & G703 Forms\Excel Waiver 1 page.xlsx"


if qtoken == 0:
    new_inv = simpledialog.askinteger("Invoice Prompt", "Enter the new invoice number:")
    completed_through= simpledialog.askinteger("Invoice Prompt", "Enter the amount billed without retention taken out:")



owner_info = bill_to.split('\r\n')
project_info = ship_to.split('\r\n')


# Page 1
workbook = load_workbook(filename=workbook_path)
sheet1 = workbook["G702"]

application = sheet1["J3"]
# Increment application number
if noe_token == 0:
    old_application_number = application.value
    print(application.value)
    application.value = application.value + 1
    new_application_number = application.value
else:
    owner = sheet1["C4"]
    owner.value = owner_info[0]
    sheet1["C5"].value = owner_info[1]
    sheet1["C6"].value = owner_info[2]
    project = sheet1["E4"]
    project.value = project_info[0]
    sheet1["E5"].value = owner_info[1]
    sheet1["E6"].value = owner_info[2]

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

# add work from this period to completed and stored to date
completed_through_cell = sheet1["E26"]
if noe_token == 0:
    completed_through_cell.value += completed_through
    old_invoice_number = invoice_number.value
    invoice_number.value = new_inv
    # add old current payment E39 to previous payments E38
    retention_percentage = sheet1['B29'].value
    sheet1["E38"].value += sheet1["E39"].value
    sheet1["E39"].value = completed_through * (1-retention_percentage*.01)
if noe_token == 1:
    completed_through_cell.value = completed_through
    invoice_number.value = new_inv
    current_payment_due = sheet1["E39"]
    current_payment_due = (completed_through * .9)



if noe_token == 0:
    # Page 2
    sheet2 = workbook["G703"]
    prev_apps = ["D13", "D17", "D21", "D25"]
    this_period = ["E13", "E17", "E21", "E25"]

    for x, y in zip(prev_apps, this_period):
        sheet2[x].value += sheet2[y].value
        sheet2[y].value = 0




# Save a new file with new name
if noe_token == 0:
    new_workbook_path = workbook_path.replace(str(old_invoice_number), str(new_inv))
    new_workbook_path = new_workbook_path.split(' ')
    new_workbook_path[-1] = new_workbook_path[-1].replace(str(old_application_number), str(new_application_number))
    new_workbook_path = (' ').join(new_workbook_path)
    print(new_workbook_path)

    workbook.save(new_workbook_path)
else:
    new_workbook_path = initial_dir + '\\' + str(owner_info[0] + ' ' + project_info[0] + ' ' + str(new_inv) + ' ' + str(application.value) + '.xlsx')
    workbook.save(new_workbook_path)
## Here I am

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
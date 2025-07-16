# If you're making a progress invoice, you'll start by manually copying the old invoice 
# TODO: Add GUI elements to create a copy of the invoice if you need it
# Potentially I could just put everything on print later for the invoices so
# i don't have to mess around with the printing interface

# The retention billing doesn't work if it's the final billing but only the second total billing. So if we bill for the total amount on invoice one, then charge for retention on invoice 2
# this doesn't work. 

# Need to add change order functionality

import pyautogui, openpyxl, datetime, calendar, os, sys, re
import subprocess, time, win32, win32com.client
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta, date

from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename, Frame, Button
from time import sleep

class OptionButtons(simpledialog.Dialog):
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

def assign_values_to_cells(values, mapping, sheet):
    for field_name, cell in mapping.items():
        if field_name in values:
            cell_value = sheet[cell].value
            if cell_value is not None:  # Ensure that the cell has a value
                try:
                    # Try to convert the cell value to float and check if it's zero or more
                    if float(cell_value) >= 0:
                        sheet[cell] = int(values[field_name])
                except ValueError:
                    # If the conversion fails, it means the value in the cell isn't a number
                    pass

def print_excel(workbook_path):
    excel = win32com.client.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(workbook_path)

    # Print the workbook to the default printer
    wb.PrintOut()

    # Clean up
    wb.Close(False)
    excel.Application.Quit()

if __name__ =="__main__":
    run_again = True
    while run_again:
        file_loc = r"C:\Users\Michael\Desktop\python-work\\" 
        # Create the Tkinter root window
        root = Tk() 
        root.withdraw()  # Hide the root window
        noe_token = 0
        new_or_existing = OptionButtons(root, title="New or Existing", button_names=["New Billing Cycle", "Existing Billing Cycle", "Retention/Final Billing"])
        print("You clicked:", new_or_existing.result)
        # change_orders =  messagebox.askyesno("Change Order", "Are there change orders?")
        # if change_orders:


        # This is a boolean that's used in progress billing to see if the retention
        # needs to be billed in order to be reduced
        retention_reduction = False

        # New means 1, existing means 0, but don't forget about Retention
        if new_or_existing.result == 'New Billing Cycle':
            noe_token = 1
        elif new_or_existing.result == "Existing Billing Cycle":
            noe_token = 0
        else:
            noe_token = "Retention"
            retention_percent = 0
            field_names = [
                "Billed Roofing Labor:", 
                "Billed Roofing Material:", 
                "Billed Sheet Metal Labor:", 
                "Billed Sheet Metal Material:",
            ]
            category_values = {field_name: 0 for field_name in field_names}

        qb_response = messagebox.askyesno("Confirmation", "Do you want to start in Quickbooks?")
        if qb_response:
            qtoken = 1
        else:
            qtoken = 2
        sleep(1)

        # New contract invoice will start using the memorized new contract invoice document
        if qb_response and noe_token == 0:
            qtoken = 1
            exec(open(f'{file_loc}progress_invoice.py').read())
        # start in QuickBooks in a new billing cycle
        elif qb_response and noe_token == 1:
        # I think this starts with you having entered the customer and job info, or you can wait to answer the "are you ready to "
        # use this billing info
            qtoken = 1
            contract_date = simpledialog.askstring("Contract Prompt", "What is the contract date MM/DD/YY:")
            exec(open(f'{file_loc}new_contract_invoice.py').read())
            owner_info = bill_to.split('\r\n')
            project_info = ship_to.split('\r\n')
        elif qb_response == 0 and noe_token == 1:
            qtoken = 0
            contract_date = simpledialog.askstring("Contract Prompt", "What is the contract date MM/DD/YY:")
            bill_to = simpledialog.askstring("Contractor", "Who is the contractor?")
            ship_to = simpledialog.askstring("Project", "What is the project name?")
        elif new_or_existing.result == "Retention/Final Billing" and qb_response:
            qtoken = 1
            exec(open(f'{file_loc}retention_invoice.py').read())
        else:
            qtoken = 0
            target_invoice = simpledialog.askinteger("Invoice number", "What is the number of the last invoice?")



        initial_dir=r"\\WBR\data\shared\G702 & G703 Forms"

        folder_path = initial_dir 


        if noe_token == 0 or noe_token == "Retention":
            old_invoice_number = target_invoice
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


        if qtoken == 0 and noe_token == "Retention":
            new_inv = simpledialog.askinteger("Invoice Prompt", "Enter the new invoice number:")
            completed_through= simpledialog.askinteger("Amount Prompt", "Enter the retention amount billed.")
        elif qtoken == 0: 
            new_inv = simpledialog.askinteger("Invoice Prompt", "Enter the new invoice number:")
            completed_through= simpledialog.askinteger("Amount Prompt", "Enter the amount billed without retention taken out:")






        # Page 1
        # This is all the header / admin info
        workbook = load_workbook(filename=workbook_path)
        sheet1 = workbook["G702"]

        application = sheet1["J3"]
        # Increment application number
        if noe_token == 0 or noe_token == "Retention":
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
            sheet1["J15"].value = contract_date


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

        # this is the actual accounting part of the page

        sheet1["B29"].value = retention_percent

        # add work from this period to completed and stored to date
        completed_through_cell = sheet1["E26"]
        retention_percentage = sheet1['B29'].value
        current_payment_due = sheet1["E39"]
        if  noe_token == "Retention" and qtoken == 0:
            retention_completed_through = completed_through
            completed_through = 0
            sheet1["E40"].value = 0
        elif noe_token == "Retention":
            completed_through = 0
            sheet1["E40"].value = 0
        if noe_token == 0 or noe_token == "Retention":
            completed_through = int(float(completed_through))
            completed_through_cell.value += int(completed_through)
            old_invoice_number = invoice_number.value
            invoice_number.value = new_inv
            # add old current payment E39 to previous payments E38
            sheet1["E38"].value += sheet1["E39"].value
            if noe_token == 0:
                if retention_reduction:
                    current_payment_due.value = int(completed_through) * (1-int(retention_percent)*.01) + (new_prev_retained / 2)
                else:
                    current_payment_due.value = int(completed_through) * (1-int(retention_percent)*.01)
            elif noe_token == "Retention" and qtoken == 1:
                current_payment_due.value = int(new_prev_retained)
            elif noe_token == "Retention" and qtoken == 0:
                current_payment_due.value = int(retention_completed_through)
        if noe_token == 1:
            # Contract
            sheet1["E23"].value = int(contract_amount)
            completed_through_cell.value = int(completed_through)
            invoice_number.value = new_inv
            current_payment_due.value = (int(completed_through) * (1-int(retention_percent)*.01))


        # Page 2

        sheet2 = workbook["G703"]

        # Excel columns
        contract_categories = ["C13", "C17", "C21", "C25"]
        prev_apps = ["D13", "D17", "D21", "D25"]
        this_period = ["E13", "E17", "E21", "E25"]

        # Mapping categaries from tk inter dialog box to excel cells
        mapping = {
            "Total Roofing Labor Contract Amount:": "C13",
            "Total Roofing Material Contract Amount:": "C17",
            "Total Sheet Metal Labor Contract Amount:": "C21",
            "Total Sheet Metal Material Contract Amount:": "C25",
            "Billed Roofing Labor:": "E13",
            "Billed Roofing Material:": "E17",
            "Billed Sheet Metal Labor:": "E21",
            "Billed Sheet Metal Material:": "E25"
        }

        if noe_token == 0 or noe_token == "Retention":
            for x, y in zip(prev_apps, this_period):
                if sheet2[x].value or sheet2[y].value:
                    sheet2[x].value += sheet2[y].value
                    sheet2[y].value = 0
        assign_values_to_cells(category_values, mapping, sheet2)


        # Save a new file with new name
        if noe_token == 0 or  noe_token == 'Retention':
            new_workbook_path = workbook_path.replace(str(old_invoice_number), str(new_inv))
            new_workbook_path = new_workbook_path.split(' ')
            new_workbook_path[-1] = new_workbook_path[-1].replace(str(old_application_number), str(new_application_number))
            new_workbook_path = (' ').join(new_workbook_path)
            print(new_workbook_path)

            workbook.save(new_workbook_path)
        else:
            new_workbook_path = initial_dir + '\\' + str(owner_info[0] + ' ' + project_info[0] + ' ' + str(new_inv) + ' ' + str(application.value) + '.xlsx')
            workbook.save(new_workbook_path)


        # TODO: Add formulae to all cells

        excel_path = r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"

        open_or_print = messagebox.askyesnocancel("Open or print", "Press yes to open file, no to print file, cancel to do neither")

        if open_or_print is True:
            # Code to open the file
            subprocess.run([excel_path, new_workbook_path])
            print('Selected open')
        elif open_or_print is False:
            # Code to print the file
            print_excel(new_workbook_path)
            print('Selected print')
        else:
            print('Selected cancel')

        print("AIA creation complete!")

        run_again = messagebox.askyesno("Continue?", "Do you want to create another AIA form?")

    scan_ready = messagebox.askyesno("Scan?", "Are you ready/do you want to start the scanning program?")
    if scan_ready:
        exec(open(f'{file_loc}AIA_scan.py').read())



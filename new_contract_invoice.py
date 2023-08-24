# I think this starts with you having entered the customer and job info, or you can wait to answer the "are you ready to "
# use this billing info

## TODO:
# I think I might be able to use tkinter to confirm when I've made a copy that way I can just automate
# the whole first part of the making a copy step and then once I've done pressed the button 
# that one time it'll do the rest for me
import pyautogui, openpyxl, time, pyperclip, datetime, calendar
import os, re, sys
import pygetwindow as gw
from openpyxl import load_workbook
from pyautogui import write, press, keyUp, keyDown, hotkey
from time import sleep
from datetime import datetime, timedelta, date

from tkinter import Tk, simpledialog, messagebox, Label, Toplevel, Button, Entry
from tkinter.filedialog import askopenfilename
from AIA import OptionButtons

def copy_clipboard():
    hotkey('ctrl', 'c')
    time.sleep(.5)
    return pyperclip.paste()

def highlight_line():
    press('numlock')
    keyDown('shiftleft')
    press('end')
    keyUp('shiftleft')
    sleep(1)
    press('numlock')
    sleep(1)

def highlight_3_line():
    hotkey('ctrl', 'home')
    press('home')
    press('numlock')
    keyDown('shiftleft')
    press('end')
    press('down', presses=2)
    press('end')
    keyUp('shiftleft')
    sleep(1)
    press('numlock')
    sleep(1)

# # def center_window(window, width, height):
#     screen_width = window.winfo_screenwidth()
#     screen_height = window.winfo_screenheight()
    
#     x = (screen_width - width) // 2
#     y = (screen_height - height) // 2
    
#     window.geometry(f"{width}x{height}+{x}+{y}")

# def show_custom_dialog():
#     dialog = Toplevel(root)
#     dialog.title("Contract Info")
    
    
#     width = 300
#     height = 300
#     center_window(dialog, width, height)

#     Label(dialog, text="Total Contract Amount:").pack()
#     contract_entry = Entry(dialog)
#     contract_entry.pack()

#     Label(dialog, text="Amount billed without retention taken out:").pack()
#     completed_entry = Entry(dialog)
#     completed_entry.pack()

#     Label(dialog, text="Total Roofing Labor Contract:").pack()
#     completed_entry = Entry(dialog)
#     completed_entry.pack()

#     Label(dialog, text="Total Roofing Material Contract:").pack()
#     completed_entry = Entry(dialog)
#     completed_entry.pack()

#     ok_button = Button(dialog, text="OK", command=dialog.destroy)
#     ok_button.pack()

#     dialog.wait_window(dialog)

#     contract_amount = contract_entry.get()
#     completed_through = completed_entry.get()

#     print("Contract Amount:", contract_amount)
#     print("Completed Through:", completed_through)

# # Create the Tkinter root window
# root = Tk()
# root.withdraw()  # Hide the root window

# show_custom_dialog()

class CustomDialog:
    def __init__(self, root, field_names):
        self.root = root
        self.dialog = Toplevel(root)
        self.dialog.title("Contract and Billing Info")
        self.field_entries = {}
        self.values = {}  # Add this line to initialize an empty dictionary for the values


        self.create_dialog(field_names)

    def center_window(self, width, height):
        self.dialog.geometry(f"{width}x{height}+200+200")


    def create_dialog(self, field_names):
        width = 300
        height = len(field_names) * 50 + 150
        self.center_window(width, height)

        for field_name in field_names:
            Label(self.dialog, text=field_name).pack()
            entry = Entry(self.dialog)
            entry.pack()
            self.field_entries[field_name] = entry

        # OK button now both fetches values and destroys the dialog
        ok_button = Button(self.dialog, text="OK", command=self.ok_clicked)
        ok_button.pack()

        self.dialog.wait_window(self.dialog)

    def get_values(self):
        return {field_name: entry.get() for field_name, entry in self.field_entries.items()}

    
    def ok_clicked(self):
        values = self.get_values()
    
        try:
            total_contract = float(values["Total Contract Amount:"])
            contract_sum = float(values["Total Roofing Labor Contract Amount:"]) + \
                        float(values["Total Roofing Material Contract Amount:"]) + \
                        float(values["Total Sheet Metal Labor Contract Amount:"]) + \
                        float(values["Total Sheet Metal Material Contract Amount:"])
        
            total_billed_period = float(values["Total Billed this period (without retention taken out):"])
            billed_sum = float(values["Billed Roofing Labor:"]) + \
                        float(values["Billed Roofing Material:"]) + \
                        float(values["Billed Sheet Metal Labor:"]) + \
                        float(values["Billed Sheet Metal Material:"])
            
            # Checking if the values match
            if not (abs(total_contract - contract_sum) < 1e-9 and abs(total_billed_period - billed_sum) < 1e-9):
                messagebox.showerror("Error", "The categories don't add up to the totals.")
                return
        except ValueError:  # Handle the case where user inputs non-numeric values
            messagebox.showerror("Error", "Please enter valid numeric values.")
            return

        # If everything is correct, then proceed
        self.values = values
        self.dialog.destroy()

def show_custom_dialog(field_names):
    root = Tk()
    root.withdraw()  # Hide the root window

    dialog = CustomDialog(root, field_names)
    category_values = dialog.values
    print(category_values)

    return dialog.values

field_names = [
    "Total Roofing Labor Contract Amount:", 
    "Total Roofing Material Contract Amount:", 
    "Total Sheet Metal Labor Contract Amount:", 
    "Total Sheet Metal Material Contract Amount:",
    "Billed Roofing Labor:", 
    "Billed Roofing Material:", 
    "Billed Sheet Metal Labor:", 
    "Billed Sheet Metal Material:",
    "Total Contract Amount:",  # New field
    "Total Billed this period (without retention taken out):",
]


category_values = show_custom_dialog(field_names)

contract_amount = category_values["Total Contract Amount:"]
completed_through = category_values["Total Billed this period (without retention taken out):"]

retention_amount_dialog = OptionButtons(Tk(), title="Retention Level", button_names=["10%", "5%"])
if retention_amount_dialog.result == '10%':
    retention_percent = 10
else:
    retention_percent = 5
print(retention_percent)

windows = gw.getAllWindows()

qb_window = None

for window in windows:
    if "QuickBooks" in window.title:
        qb_window = window
        break


invoice_status = messagebox.askyesno("Confirmation", "Have you already created the invoice?")

# contract_amount = simpledialog.askinteger("Contract Prompt", "What is the total contract amount:")

# completed_through= simpledialog.askinteger("Invoice Prompt", "Enter the amount billed without retention taken out:")

qb_window.activate()

# If you haven't started an invoice yet, this will create a contract invoice
if not invoice_status:
    hotkey('ctrl', 't')
    time.sleep(.5)
    press('up', presses=25)
    press('down', presses=3)
    hotkey('alt', 't')
    sleep(2)

bill_bin = messagebox.askyesno("Confirmation", "Are you ready to have this invoice information used?")
sleep(1)
qb_window.activate()
sleep(.5)
hotkey('alt', 'j')
time.sleep(.5)

sleep(1)

today = date.today()
print(today)
res = calendar.monthrange(today.year, today.month)[1]
completed_date = f"Completed through {today.month}/{res}/{today.year}"



# 8 tabs to the first item
# 10 tabs to the first price cell

# get the new invoice number
press('tab', presses=3)
new_inv = copy_clipboard()

sleep(1)

press('tab')
highlight_3_line()
bill_to = copy_clipboard()
sleep(1)

press('tab', presses=2)
highlight_3_line()
ship_to = copy_clipboard()
sleep(.5)


press('tab', presses=4)
write(str(contract_amount))
press('down', presses=1)
sleep(.5)
write(str(completed_through))
sleep(.5)
hotkey('shift', 'tab')


highlight_line()

press('backspace')
write(completed_date)
press('tab')

press('down')
write(f'-{retention_percent}%')
press('tab')


print_bin = messagebox.askyesno("Confirmation", "Do you want to print this?")

if print_bin:
    sleep(1)
    qb_window.activate()
    hotkey('ctrl', 'p')
    sleep(5)
    press('space') 
else:
    sys.exit()


# This is some code from ChatGPT that lets you find a file by the number
# even if it's not the only thing in the file name


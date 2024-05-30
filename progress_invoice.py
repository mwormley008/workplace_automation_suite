## TODO:

# I need to make a first progress invoice because if it's the first time
# you need to add additional cells that aren't in the existing multiple billing
# cycle version

import pyautogui, openpyxl, time, pyperclip, datetime, calendar
import os, re, sys
import pygetwindow as gw
from openpyxl import load_workbook
from pyautogui import write, press, keyUp, keyDown, hotkey
from time import sleep
from datetime import datetime, timedelta, date

from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename
from AIA import OptionButtons
from click_qb_copy_button import click_duplicate_button
from new_contract_invoice import CustomDialog, show_custom_dialog


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

def highlight_multi_line(number_of_lines):
    press('numlock')
    keyDown('shiftleft')
    press('end')
    press('down', presses=number_of_lines)
    keyUp('shiftleft')
    sleep(1)
    press('numlock')
    sleep(1)

def find_invoice(invoice_number):
    hotkey('ctrl', 'f')
    press('tab', presses=4)
    write(str(invoice_number))
    press('enter')
    sleep(1.5)
    hotkey('alt', 'g')
    sleep(1)

def bill_reduction(retention_to_reduce):
    sleep(.5)
    press('tab')
    write("New Ret Due")
    press('tab')
    sleep(.5)
    press('tab')
    sleep(.5)
    retention_to_reduce /= 2
    retention_to_reduce *= -1
    write(str(retention_to_reduce))

def number_formatting(number):
    number = number.replace(',', '')
    number = number[0:-3]
    number = int(number)
    return number
    

    
def new_ok_clicked(self):
        values = self.get_values()
    
        try:
            
            total_billed_period = float(values["Total Billed this period (without retention taken out):"])
            billed_sum = float(values["Billed Roofing Labor:"]) + \
                        float(values["Billed Roofing Material:"]) + \
                        float(values["Billed Sheet Metal Labor:"]) + \
                        float(values["Billed Sheet Metal Material:"])
            
            # Checking if the values match
            if not (abs(total_billed_period - billed_sum) < 1e-9):
                messagebox.askyesno("Error", "The categories don't add up to the totals, do you want to proceed anyway?")
        except ValueError:  # Handle the case where user inputs non-numeric values
            messagebox.showerror("Error", "Please enter valid numeric values.")
            return

        # If everything is correct, then proceed
        self.values = values
        self.dialog.destroy()

def change_orders():
    cho_amt = simpledialog.askinteger("Change orders", "How many change orders does this invoice have?")
    return cho_amt

if __name__ == "__main__":
    APP_PATH = r"C:\Program Files (x86)\Intuit\QuickBooks 2019\QBW32.EXE"
    MAIN_WINDOW_TITLE_REGEX = ".*QuickBooks Desktop Pro 2019.*"
    CHILD_WINDOW_TITLE = "Create Invoices (Editing Transaction...) "
    BUTTON_ID = "DuplicateBtn"
    
    
    CustomDialog.ok_clicked = new_ok_clicked
    
    
    
    windows = gw.getAllWindows()

    qb_window = None



    for window in windows:
        if "QuickBooks" in window.title:
            qb_window = window
            break

    # Create the Tkinter root window
    root = Tk()
    root.withdraw()  # Hide the root window

    second_bill = messagebox.askyesno("Billing Lifecycle", "Is this the second entry in the billing cycle?")

    field_names = [
            "Number of previous invoice",
            "Total Billed this period (without retention taken out):",
            "Billed Roofing Labor:", 
            "Billed Roofing Material:", 
            "Billed Sheet Metal Labor:", 
            "Billed Sheet Metal Material:",
        ]

    category_values = show_custom_dialog(field_names)

    target_invoice = category_values["Number of previous invoice"]

    completed_through = category_values["Total Billed this period (without retention taken out):"]
    retention_amount_dialog = OptionButtons(root, title="Retention Level", button_names=["10%", "5%"])

    if retention_amount_dialog.result == '10%':
        retention_percent = 10
    else:
        retention_percent = 5
        retention_reduction = messagebox.askyesno("Reduction status", "Bill to reduce retention from 10% to 5%")
        sleep(1)
    print(retention_percent)

    cho_num = change_orders()

    # Focuses the Quickbooks window and goes to the customer:job pane
    qb_window.activate()

    find_invoice(target_invoice)
    
    sleep(1)

    click_duplicate_button(APP_PATH, MAIN_WINDOW_TITLE_REGEX, CHILD_WINDOW_TITLE, BUTTON_ID)

    sleep(1)

    press('space')

    sleep(.3)

    hotkey('alt', 'j')
    time.sleep(.5)

    sleep(1)

    today = date.today()
    print(today)
    res = calendar.monthrange(today.year, today.month)[1]
    completed_date = f"Completed through {today.month}/{res}/{today.year}"


    # starts once you have created a copy of the invoice, which will
    # start with highlighting the customer job
    # 8 tabs to the first item
    # 10 tabs to the first price cell

    # get the new invoice number
    press('tab', presses=3)
    new_inv = copy_clipboard()

    # gets the previously billed
    if cho_num == 0:
        nca = 0
        # Goes to the contract amount / the first cell of the price column
        press('tab', presses=7)
        # This is previously billed
        press('down', presses=1)
        sleep(.5)
        prev_billed = copy_clipboard()
        sleep(.5)
        prev_billed = prev_billed.replace(',', '')
        prev_billed = prev_billed[0:-3]
    else:
        nca = 1
        # Goes to the contract amount / the first cell of the price column
        press('tab', presses=7)
        # This is previously billed
        press('down', presses=2+cho_num)
        sleep(.5)
        prev_billed = copy_clipboard()
        sleep(.5)
        prev_billed = prev_billed.replace(',', '')
        prev_billed = prev_billed[0:-3]
    # gets total retention and billed last period and last period's retention
    if not second_bill:
        press('down')
        new_prev_retained = copy_clipboard()
        new_prev_retained = number_formatting(new_prev_retained)
        # Goes to the last invoice's completed through which will become the last perioud amount
        press('down')
        last_period = copy_clipboard()
        sleep(.5)

        # This could become a method or something since I'm doing it multiple times and it would improve 
        # Legibility
        last_period = last_period.replace(',', '')
        last_period = last_period[0:-3]
        press('down')
        press('tab')
        new_prev_retained -= number_formatting(copy_clipboard())
        # number_formatting(new_prev_retained)
    else:
        press('down')
        press('tab')
        new_prev_retained = number_formatting(copy_clipboard())
        last_period = 0

    hotkey('shift', 'tab')
    press('up')

    # Combines the previously billed with the amount billed last perioud for the new previously billed
    new_prev_billed = int((prev_billed)) + int((last_period))

    

    # Alright, at this point we now have all of the information and now we're just going to be writing it to
    # The quickbooks fields

    # This starts in the current period cell, which is higher up in the second billing than others
    if not second_bill:
        press('up', presses=2)
        sleep(.5)
        press('backspace')
        write(str(new_prev_billed))
        press('down')
        press('backspace')
        write(str(new_prev_retained))
        press('tab', presses=4)
        highlight_line()
        # sleep(2)
        press('backspace')
        write(completed_date)
        press('tab')
        write(str(completed_through))
        press('down')
        write(f'-{retention_percent}%')
        press('tab')
        if retention_reduction:
            bill_reduction(new_prev_retained)

            # completed_through = str(int(completed_through) + int(new_prev_retained))
    else:
        press('tab')
        write('0')
        hotkey('shift', 'tab')
        hotkey('shift', 'tab')
        highlight_line()
        write('Previously billed')
        hotkey('shift', 'tab')
        write('PREV')
        press('down')
        sleep(1)
        write('PREV RET')
        press('tab')
        highlight_line()
        write("Previously retained")
        press('tab')
        write(str(new_prev_retained))
        press('tab')
        write('0')
        press('tab', presses=2)
        write('L&M')
        press('tab')
        highlight_multi_line(2)
        press('backspace')
        write(completed_date)
        # sleep(2)
        press('tab')
        write(str(completed_through))
        sleep(.5)
        sleep(.5)
        press('tab')
        write('1')
        press('tab', presses=2)
        write('Retention')
        sleep(.5)
        press('tab', presses=2)
        write(f'-{retention_percent}%')
        press('tab')
        if retention_reduction:
            bill_reduction(new_prev_retained)
            # completed_through = str(int(completed_through) + int(new_prev_retained))
        

    press('tab')
    sleep(.5)
    print_bin = messagebox.askyesno("Confirmation", "Do you want to print this?")

    if print_bin:
        sleep(1)
        qb_window.activate()
        hotkey('ctrl', 'p')
        # sleep(5)
        # press('space') 
    else:
        sys.exit()


    # This is some code from ChatGPT that lets you find a file by the number
    # even if it's not the only thing in the file name


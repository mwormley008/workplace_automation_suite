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
completed_through= simpledialog.askinteger("Invoice Prompt", "Enter the amount billed without retention taken out:")


retention_amount_dialog = OptionButtons(Tk(), title="Retention Level", button_names=["10%", "5%"])
if retention_amount_dialog.result == '10%':
    retention_percent = 10
else:
    retention_percent = 5
print(retention_percent)


# Focuses the Quickbooks window and goes to the customer:job pane
qb_window.activate()
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


# Goes to the contract amount / the first cell of the price column
press('tab', presses=7)
# This is previously billed
press('down', presses=1)
sleep(.5)
prev_billed = copy_clipboard()
sleep(.5)
prev_billed = prev_billed.replace(',', '')
prev_billed = prev_billed[0:-3]

if not second_bill:
    # Goes to the last invoice's completed through which will become the last perioud amount
    press('down', presses=2)
    last_period = copy_clipboard()
    sleep(.5)

    # This could become a method or something since I'm doing it multiple times and it would improve 
    # Legibility
    last_period = last_period.replace(',', '')
    last_period = last_period[0:-3]
else:
    last_period = 0

# Combines the previously billed with the amount billed last perioud for the new previously billed
new_prev_billed = int((prev_billed)) + int((last_period))

# Sets previously retained
new_prev_retained = new_prev_billed * .1

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
    write('PREV RET')
    press('tab', presses=2)
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


# Auto GUI
# This is a program that grabs information from a specified 
# Excel column and then automates the input of that list into a data entry form
# using autogui



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

from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename

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

invoice_status = messagebox.askyesno("Confirmation", "Have you already created the invoice?")

contract_amount = simpledialog.askinteger("Contract Prompt", "What is the total contract amount:")

completed_through= simpledialog.askinteger("Invoice Prompt", "Enter the amount billed without retention taken out:")

qb_window.activate()

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


# starts once you have created a copy of the invoice, which will
# start with highlighting the customer job
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
write('-10%')


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


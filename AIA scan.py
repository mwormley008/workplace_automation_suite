# Script to save a scan and email it once the invoice is scanned, then you just need the
# invoice number, that's it
# you'll need your scan window still open as well as a chrome window
# in the wbr gmail tab

# Scan your invoice then begin


# TODO: I think I'd like to make it so that if you don't have the WBR 
# gmail open it'll open it for you although that might be solvable with the 
# gmail api

import pyautogui, openpyxl, datetime, calendar, os, sys, re
import subprocess
import pygetwindow as gw
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta, date

from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename
from pyautogui import press, write, hotkey
from time import sleep

from download_repair_photos import send_message

# This is to do things relatively, this one would be to do things like click the make a copy button in QB or the scan button in scansmart
"""
import pyautogui

# Find the window position
window_x, window_y, window_width, window_height = pyautogui.locateOnScreen('window.png')

# Calculate the button position relative to the window
button_x = window_x + button_relative_x
button_y = window_y + button_relative_y

# Move the mouse cursor to the button position and click
pyautogui.moveTo(button_x, button_y)
pyautogui.click()
"""



mail_contacts = {
    'Novak':'DCaporale@novakconstruction.com', 
    'Valenti':'Melissa.Sanson@valentibuilders.com', 
    'Hanna':'mrosales@hannadesigngroup.com',
    'G&H':'theresa@nationalplazas.com',
    'Builtech':'dwiniarz@builtechllc.com'}
# Create the Tkinter root window
root = Tk()
root.withdraw()  # Hide the root window


windows = gw.getAllWindows()

scan_window = None
wbr_window = None

for window in windows:
    if "ScanSmart" in window.title:
        scan_window = window
    if "wbrroof@gmail.com" in window.title:
        wbr_window = window

sleep(1)

invoice_number = simpledialog.askinteger("Invoice Prompt", "Enter the invoice number:")

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
target_number = invoice_number  # Replace with the specific number you are looking for

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

file_name = workbook_path

file_name = file_name.replace(initial_dir + "\\", "")
file_name = file_name[0:-5]
print(file_name)

scan_window.activate()
pyautogui.press('enter')
sleep(1)
pyautogui.write(file_name)
sleep(2)
press('enter')

wbr_window.activate()
sleep(1)

press('esc')
press('c')
sleep(2)
company = file_name.split(' ')[0]
if company in mail_contacts:
    write(mail_contacts[company])
sleep(1)
press('tab')
sleep(.5)
press('tab')
sleep(.5)
write(file_name)
sleep(1)
press('tab')
write('Hello,\nPlease see attached billing.\nThank you,\nMichael Wormley\nWBR Roofing\n25084 W Old Rand Rd\nWauconda, IL 60084\n​O: 847-487-8787​\nwbrroof@aol.com')
sleep(3)
press('tab', presses=3)
press('space')
sleep(2.5)
write(file_name)
sleep(1)
press('down')
press('enter')
sleep(2)
hotkey('ctrl', 'enter')

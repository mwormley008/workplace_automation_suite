# Inputs invoices from ABC myabcsupply, with the open bills window open in QB

import pyautogui, openpyxl, datetime, calendar, os, sys, re
import subprocess, time
import pygetwindow as gw
import csv

from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta, date

from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename, Frame, Button
from time import sleep
from pyautogui import press, write, hotkey



file_name = f'ABC.csv'
folder_path = r'C:\Users\Michael\Desktop\python-work\Invoices'

# invoices_for_input = load_workbook(os.path.join(folder_path, file_name))

def quick_windows():
    windows = gw.getAllWindows()

    qb_window = None

    for window in windows:
        if "QuickBooks Desktop" in window.title:
            qb_window = window

    return qb_window        

# ws = invoices_for_input['Sheet']

first_row_skipped = False

# Alright, this starts with the bill window open in QB

qb_window = quick_windows()
sleep(1)

# Focuses QuickBooks window
if qb_window:
    # pass
    qb_window.activate()
else:
    print("QuickBooks window not found.")   
sleep(3)



with open(os.path.join(folder_path, file_name), 'r', newline='', encoding='utf-8') as file:
    reader = csv.reader(file)
    next(reader)  # skip the header row

    for row in reader:
        col_A, col_B, col_C, col_D, col_E, col_F, col_G, col_H  = row
        print(col_A, col_B, col_C, col_D, col_E, col_F, col_G, col_H)
        write('ABC Supply')
        sleep(1)
        pyautogui.press('tab')
        sleep(1.5)
        pyautogui.write(col_B)
        sleep(1.5)
        pyautogui.press('tab')
        sleep(1.5)
        pyautogui.write(col_A)
        sleep(1)
        pyautogui.press('tab')
        sleep(.5)
        pyautogui.write(col_C)
        sleep(.5)
        # This part is new to tag the invoice entry
        pyautogui.press('tab', presses=3)
        sleep(.5)
        write(f'{col_D}')
        sleep(.5)
        hotkey('alt', 's')
        sleep(2)

print('Entered!')


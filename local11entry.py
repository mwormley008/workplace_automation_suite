# Auto GUI
# This is a program that grabs information from a specified 
# Excel column and then automates the input of that list into a data entry form
# using autogui 

## You'll need to log in to the union website, copy the old 
## report, then start in the AH column
import pyautogui, openpyxl, time
from openpyxl import load_workbook
from pyautogui import write, press
from time import sleep
from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename
import pygetwindow as gw


# Select a file from the folder where the files are
# sleep(3)
initial_dir = r"\\WBR\shared\PAYROLL\WBR Payroll WE 2022"
workbook_path = askopenfilename(initialdir=initial_dir)


# For some reason this block isn't working
# The if statement is getting met but when the window tries to activate it doesn't work

# windows = gw.getAllWindows()
# for window in windows:
#     if "employer.gobasys" in window.title:
#         window.activate()


sleep(1)

# workbook_path = r"\\WBR\shared\PAYROLL\WBR Payroll WE 2022\Local 11 June 2023.xlsx"

workbook = load_workbook(filename=workbook_path, data_only=True)

sleep(1)

# Sets the worksheet and column (this is correct without moving anything)
sheet = workbook["Sheet1"]
column_range = sheet["I5:I16"]

# 
column_list = []
for cell in column_range:
    column_list.append(cell[0].value or "")

# print(column_list)

for i in column_list:
    # print(i)
    sleep(.1)
    write(str(i))
    sleep(.5)
    press('tab')
    sleep(.4)
# Auto GUI
# This is a program that grabs information from a specified 
# Excel column and then automates the input of that list into a data entry form
# using autogui 
import pyautogui, openpyxl, time
from openpyxl import load_workbook
from pyautogui import write, press
from time import sleep

sleep(3)

workbook_path = r"\\WBR\shared\PAYROLL\WBR Payroll WE 2022\Local 11 June 2023.xlsx"
workbook_path = r"\\WBR\shared\PAYROLL\WBR Payroll WE 2022\Local 11 June 2023.xlsx"

workbook = load_workbook(filename=workbook_path, data_only=True)

sheet = workbook["Sheet1"]
column_range = sheet["I5:I16"]

column_list = []
for cell in column_range:
    column_list.append(cell[0].value or "")

# print(column_list)

for i in column_list:
    # print(i)
    write(str(i))
    press('tab')
# Script to put in the expense categories

import pyautogui, openpyxl
from openpyxl import Workbook, load_workbook



workbook_path = r"\\WBR\data\shared\G702 & G703 Forms\41 North Contractors Raising Canes New Lenox 15682 3.xlsx"
"\\WBR\data\shared\G702 & G703 Forms\41 North Contractors Raising Canes New Lenox 15682 3.xlsx"
# Open AIA form based on invoice number
# maybe I could do this with tkinter idk
## If I have tkinter form I guess it could prompt old inv number,
## new invoice number, and total completed and stored to date

# Page 1
workbook = load_workbook(filename=workbook_path)
sheet = workbook["G702"]
application = sheet.cell("J3")
print(application)
# Increment application number
# cell = j3



# increment work period
# cell j7

# change invoice number to the new invoice number
# J9

# add old current payment E39 to previous payments E38


# Page 2

# Add old this period to previous applications
#D13 E13
# Each new row increases by 4 (D17 E17)

# Save a new file with new name

# Add formulae to all cells 
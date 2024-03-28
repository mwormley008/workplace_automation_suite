from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename

file_loc = r"C:\Users\Michael\Desktop\python-work\\" 
bill_status = messagebox.askyesno("Scan branch", "Press yes to scan proposals, press no to scan AIA forms")

if bill_status:
    exec(open(f'{file_loc}proposal_scan.py').read())
else:
    exec(open(f'{file_loc}AIA_scan.py').read())

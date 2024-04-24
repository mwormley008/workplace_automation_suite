# Script to save a scan and email it once the invoice is scanned, then you just need the
# invoice number, that's it

## This starts with the invoice unscanned in the scanner, but with the scanner program started



# TODO:
""" also I can add a branching path so that I can run one program to scan for either billing 
or for proposals

alright so I can create a database so that I can associate company's with their accountant
or their estimator 

also I can load gmail contacts from google people in order to ask for better default options"""

import pyautogui, openpyxl, datetime, calendar, os, sys, re
import subprocess
import pygetwindow as gw
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta, date

from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename
from pyautogui import press, write, hotkey
from time import sleep
from pywinauto import Application, timings

from photos_timesheets import send_message
from AIA_scan import find_or_run_scan
from send_gmail import initialize_service, create_message, create_message_with_attachment, send_message, CLIENT_SECRET_FILE, SCOPES, API_NAME, API_VERSION

def find_file_by_number(folder_path, target_number):
    pattern = re.compile(r'.*{}.*\.xlsx$'.format(target_number))

    for filename in os.listdir(folder_path):
        if pattern.match(filename):
            file_path = os.path.join(folder_path, filename)
            return file_path

    # If the file is not found
    return None

def auto_manual_email_from_aia_scan():
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
    write(proposal_message)
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

def find_scan_and_email_windows():
    windows = gw.getAllWindows()

    scan_window = None
    wbr_window = None

    for window in windows:
        if "ScanSmart" in window.title:
            scan_window = window
        
    return scan_window

def trim_extension_format_slashes(path, initial_dir):
    file_name = path.replace(initial_dir.replace("\\", "/"), "")
    file_name = file_name[1:-4]
    if file_name.endswith("."):
        file_name = file_name[:-1]
    return file_name

def click_scan_button():
    # This uses pyautogui to locate the scan button using a png because pywinauto was taking over a minute and
    # and a half to click the button once it had found it
    start = pyautogui.locateCenterOnScreen('scansingle.png')#If the file is not a png file it will not work
    print(start)
    pyautogui.moveTo(start)#Moves the mouse to the coordinates of the image
    pyautogui.click()

def click_save_button():
    # This uses pyautogui to locate the scan save button using a png because pywinauto was taking over a minute and
    # and a half to click the button once it had found it
    start = pyautogui.locateCenterOnScreen('savescan.png')#If the file is not a png file it will not work
    print(start)
    pyautogui.moveTo(start)#Moves the mouse to the coordinates of the image
    pyautogui.click()



if __name__ == '__main__':
    run_again = True
    timings.Timings.fast()
    app = Application(backend="uia").start(r"C:\Program Files (x86)\Epson Software\Epson ScanSmart\ScanSmart.exe")
    while run_again == True:
            
        find_or_run_scan()

        mail_contacts = {
            # 'Novak':'DCaporale@novakconstruction.com', 
            # 'Valenti':'billings@valentibuilders.com', 
            # 'Hanna':'mrosales@hannadesigngroup.com',
            # 'G&H':'theresa@nationalplazas.com',
            # 'Builtech':'dwiniarz@builtechllc.com',
            # 'Englewood':'VLara@eci.build',
            # 'Ott':'kate@ottdevelopment.com',
            # 'R.':'nickc@rcarlsonandsons.com',
            # '41':'amy.hillgamyer@41northcontractors.com'
            'Stasica':'todd@stasicaconstruction.com',
            'LBX':'danielle@lbxinvestments.com',
            'Raby':'daveraby@rabyroofing.com',
            'Crossroads':'christine@xr-partners.com',
            'Stratus':'David.Rekstad@stratusunlimited.com',
            'Raffin':'albertr@raffinconstruction.com',
            'ACT':'nmorrison@actconstruction.com',
            'Quincy':'twiersma@quincyconstruction.com',
            'Allstar':'allstar@allstardentalinc.com',
            'Gordon':'kelly@midwestearnosethroat.com',
            'Favia':'mike@faviabuilders.com',
            'G.A.':'jgalbraith@gajohnson.com; igalbraith@gajohnson.com',
            'Mariottini':'jeffm@mariottini.com',
            'Modell':'modelfh@hotmail.com',
            'Mosaic':'dd@mosaic-pd.com',
            'Montemayor':'joe@montemayorconstruction.com',
            'Tessler':'kurt@tesslerinc.com; ron@tesslerinc.com',
            'test':'wbrroof@gmail.com; throwod@gmail.com',
            'Tim':'briana@timcoteinc.com',
            'J.':'j jdivita.associates@yahoo.com',
            }

        proposal_message = 'Hello,\nPlease see attached proposal.\n\nThank you,\nMichael Wormley\nWBR Roofing\n25084 W Old Rand Rd\nWauconda, IL 60084\n​O: 847-487-8787​\nwbrroof@aol.com'
        # Create the Tkinter root window
        root = Tk()
        root.withdraw()  # Hide the root window


        sleep(1)


        initial_dir=r"\\WBR\data\shared\Proposals"
        initial_scan_dir=r"\\WBR\data\shared\My Scans"

        proposal_path = askopenfilename(initialdir=initial_dir)
        
        
        scan_window = find_scan_and_email_windows()
        print(proposal_path)
        
        # file_name = proposal_path
        # file_name = proposal_path.replace(initial_dir.replace("\\", "/"), "")
        file_name = trim_extension_format_slashes(proposal_path, initial_dir)
        
        # file_name = file_name.replace('/' + initial_dir, "")
        # file_name = file_name[1:-4]
        print(file_name)
        if len(file_name) > 63:
            long_file_name = True
        else:
            long_file_name = False
        print(long_file_name)
        
        
        scan_window.activate()
        
        #TODO THIS IS Really really slow now for some reason
        # Connect to the application (you might need to adjust this part based on your application details)
        # app = Application(backend="uia").connect(title="Epson ScanSmart")
        print(datetime.now())
        # Navigate to the button using its AutomationId
        scan_button = app.window(title="Epson ScanSmart").child_window(auto_id="SingleSidedScanButton")
        print(scan_button)
        

        print(datetime.now())
        # scan_button = app.child_window(auto_id="SingleSidedScanButton")
        click_scan_button()

        print(datetime.now())
        
        sleep(1)
        # # Invoke the button
        # continue_button = messagebox.askyesno("Did you click the scan button?")
        # sleep(.5)

        while True:
            try:
                if (pyautogui.locateOnScreen('savescan.png') is not None):
                    print("Save button found!")
                    click_save_button()
                    break
                else:
                    print("Not yet lol")

            except Exception as e:
                # If an error occurs (like the button isn't found), wait for 2 seconds and try again.
                sleep(2)
        sleep(.5)
        press('enter')
        sleep(.5)
        pyautogui.write(file_name)
        sleep(2)
        press('enter')
        sleep(1)

        ## If the file name is too long, it may be a duplicate because the ending will be cut off 
        if long_file_name:
            proceed = messagebox.askyesno("Proceed?", "Are you ready to proceed? Your file name was long, so you may have needed a custom file name.")

        company = file_name.split(' ')[0]
        if company in mail_contacts:
            recipient = mail_contacts[company]
        else:
            recipient = simpledialog.askstring("No default email found", f"Enter the email address for estimator at {company}:")
            with open('estimators.txt', 'a') as file:
                file.write(f"\n'{company}':'{recipient}',")


        print(recipient)

        service = initialize_service()
        if long_file_name:
            scan_path = askopenfilename(initialdir=initial_scan_dir)
            file_name = trim_extension_format_slashes(scan_path, initial_scan_dir)
        
        
        message = create_message_with_attachment(recipient, file_name, proposal_message, rf'\\WBR\data\shared\My Scans\{file_name}.pdf')

        send_message(service, 'me', message)


        run_again = messagebox.askyesno("Continue?", "Do you want to scan another proposal?")
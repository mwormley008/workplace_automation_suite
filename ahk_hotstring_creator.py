# Program to take highlighted word to append to hotstrings, plus prompt for the shortcut you want to use
# Also to restart ahk

from progress_invoice import copy_clipboard
import tkinter as tk
from tkinter import simpledialog
import subprocess
import time, os
from time import sleep


def get_user_input(text):
    # this prompts the user to provide a trigger text, I'm using it in the add_hotstring
    user_string = simpledialog.askstring("Input", f"Please enter the trigger text for '{text}'")
    
    # Replace the print statement with whatever you want to do with the input
    if user_string:
        print(user_string)
    return user_string



def add_hotstring():
    # this takes a highlighted piece of text, then prompts user to provide trigger text and then 
    # appends that to the ahk script
    replacement_text = copy_clipboard()
    trigger_text = get_user_input(replacement_text)
    with open(ahk_script, "a") as file:
        file.write(f"\n:*:{trigger_text}/::{replacement_text}")
    print("Added hotstring.")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide the root window


    ahk_script = r"C:\Users\Michael\Desktop\Text Replacement.ahk"
    ahk_executable = r"C:\Program Files\AutoHotkey\AutoHotkey.exe"  # Typical path, adjust if different
    bat_file_path = r"C:\Users\Michael\Desktop\python-work\run_ahk.bat"


    add_hotstring()
    task_name = "StartMyAHKScript"
    subprocess.run(["schtasks", "/run", "/tn", task_name])
    # subprocess.Popen([ahk_executable, ahk_script], creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP)

    # subprocess.Popen([bat_file_path], creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP)

    time.sleep(1)





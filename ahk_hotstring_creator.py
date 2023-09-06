# Program to take highlighted word to append to hotstrings, plus prompt for the shortcut you want to use
# Also to restart ahk

from progress_invoice import copy_clipboard
import tkinter as tk
from tkinter import simpledialog

def get_user_input(text):
    # this prompts the user to provide a trigger text, I'm using it in the add_hotstring
    user_string = simpledialog.askstring("Input", f"Please enter the trigger text for '{text}'")
    
    # Replace the print statement with whatever you want to do with the input
    if user_string:
        print(user_string)


def add_hotstring():
    # this takes a highlighted piece of text, then prompts user to provide trigger text and then 
    # appends that to the ahk script
    replacement_text = copy_clipboard()
    trigger_text = get_user_input(replacement_text)
    with open(ahk_script, "a") as file:
        file.write(f"\n :*:{trigger_text}/::{replacement_text}\n")
    print("Added hotstring.")

root = tk.Tk()
root.withdraw()  # Hide the root window


ahk_script = r"C:\Users\Michael\Desktop\Text Replacement.ahk"

add_hotstring()


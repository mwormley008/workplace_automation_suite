# forward emails from aol to gmail to print

import pyautogui as pya, pyperclip, time, re

from pyautogui import press, write, hold

time.sleep(3)

def copy_clipboard():
    pya.hotkey('ctrl', 'c')
    time.sleep(.01)
    return pyperclip.paste()

def highlight_line():
    pya.press('numlock')
    pya.keyDown('shiftleft')
    pya.press('home')
    pya.keyUp('shiftleft')
    pya.press('numlock')

contents = []

address = 'wbrroof@gmail.com'

highlight_line()

def aol_forward():
    pya.press('tab', presses=2)

pya.write(str(total))


pya.hotkey('shift', 'tab')
pya.press('down')
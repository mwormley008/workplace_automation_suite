# men and hours

import pyautogui as pya, pyperclip, time

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

#contents = []

highlight_line()

full_string = copy_clipboard()

','.split(full_string)
print(full_string)


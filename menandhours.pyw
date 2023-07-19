# men and hours

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

highlight_line()

full_string = copy_clipboard()

full_string = full_string.split(', ')
print(full_string)

for i in full_string:
    contents.extend(i.split())

print(contents)

total = int(contents[0]) * float(contents[2])
print(total)

pya.press('tab', presses=2)
pya.write(str(total))

pya.hotkey('shift', 'tab')
pya.press('down')
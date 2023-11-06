# men and hours

import pyautogui as pya, pyperclip, time, re

from time import sleep

from pyautogui import press, write, hold

import pygetwindow as gw
windows = gw.getAllWindows()

# sleep(.1)
qb_window = None

for window in windows:
    if "QuickBooks" in window.title:
        qb_window = window
        break

qb_window.activate()

# time.sleep(3)

def copy_clipboard():
    pya.hotkey('ctrl', 'c')
    # time.sleep(.01)
    return pyperclip.paste()

def highlight_line():
    pya.press('numlock')
    pya.keyDown('shiftleft')
    pya.press('home')
    pya.keyUp('shiftleft')
    pya.press('numlock')

contents = []


highlight_line()
# sleep(.5)

copy_clipboard()

full_string = copy_clipboard()

full_string = full_string.split('m')
# print(full_string)
pya.write(f'{full_string[0]} men, {full_string[1]} hours')

# for i in full_string:
#     contents.extend(i.split())

# print(contents)

total = int(full_string[0]) * float(full_string[1])
# print(total)

pya.press('tab', presses=2)
# sleep(.5)
pya.write(str(total))
# sleep(.5)

pya.hotkey('shift', 'tab')
pya.press('down')
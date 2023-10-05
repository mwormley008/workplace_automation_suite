from pyautogui import press, write, hold, hotkey
from time import sleep

print("Starting the script...")
sleep(2)

print("Opening Find and Replace dialog...")
hotkey('ctrl', 'h')
sleep(0.5)

print("Entering ( character...")
press('(')
sleep(0.1)

print("Pressing Alt+A...")
hotkey('alt', 'a')
sleep(0.1)

print("Entering replacement...")
press('space') 
sleep(0.1)
press(')')
sleep(0.1)
hotkey('alt', 'a')
sleep(0.1)
press('space')

print("Exiting...")
press('esc')

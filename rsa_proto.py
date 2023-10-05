import pyautogui, time
from pyautogui import write, press, hotkey
from time import sleep

def sleep_tab(number_of_tabs, len_of_sleeps):
        for i in number_of_tabs:
            press('tab')
            sleep(len_of_sleeps)
            
def paste():
    hotkey('ctrl', 'v')

def chase():
    paste()
    sleep_tab(1, 3)
    
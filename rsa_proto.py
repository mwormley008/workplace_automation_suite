import pyautogui, time
from pyautogui import write, press, hotkey
from time import sleep

def sleep_tab(number_of_tabs, len_of_sleeps):
        for i in range(number_of_tabs):
            press('tab')
            sleep(len_of_sleeps)
            
def paste():
    hotkey('ctrl', 'v')

def chase_buy():
    paste()
    sleep_tab(1, 5)
    press('tab', presses=3)
    press('space')
    press('tab')
    press('space')
    press('tab')
    write('1')
    press('tab')
    press('space')

def schwab_buy():
    # For this one I can't get the enter button to actually get pushed
    # and for sme reason the number of tabs is varible depending on the stock? 
    # or the number order that it is?
    # paste()
    # sleep(5)
    # pyautogui.press('return')
    # sleep(1)
    # sleep_tab(1, 1)
    sleep_tab(8, .3)
    press('space')
    press('b')
    sleep_tab(5, .3)
    press('m', presses=2)



# chase_buy()
schwab_buy()
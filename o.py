import pyautogui, pyperclip
import time
from pyautogui import press, write, hotkey, keyDown, keyUp
from time import sleep

# Define a helper function to send keys with a delay
def send_keys_with_delay(*keys, delay=0.01):
    pyautogui.press(keys, interval=delay)

def repeat_hotkey(times, delay, *keys):
    for _ in range(times):
        pyautogui.hotkey(*keys)
        sleep(delay)

def copy_clipboard():
    hotkey('ctrl', 'c')
    time.sleep(.5)
    return pyperclip.paste()

def highlight_line():
    press('numlock')
    keyDown('shiftleft')
    press('end')
    keyUp('shiftleft')
    sleep(1)
    press('numlock')
    sleep(1)

def highlight_home_line():
    press('numlock')
    keyDown('shiftleft')
    press('home')
    keyUp('shiftleft')
    sleep(1)
    press('numlock')
    sleep(1)

def search_term(term):
    pyautogui.hotkey('ctrl', 'f')
    sleep(1)
    pyautogui.write(term)
    sleep(1)
    pyautogui.press('return')
    sleep(1)
    pyautogui.press('esc')
    sleep(1)

## Add a focus function for .doc files


# Start the flow
time.sleep(2)  # WAIT 2



def NameSave():
    # Search for "proposal submitted to:"
    search_term('proposal submitted to:')
    send_keys_with_delay('down', 'end')
    highlight_line()
    contractor = copy_clipboard()
    

    # Search for "Work to be performed at:"
    search_term('Work to be performed at:')
    send_keys_with_delay('down', 'end', delay=0.14)
    highlight_line()
    job_name = copy_clipboard()
    sleep(1)
    

    # Search for "Work to be performed at:" again
    
    send_keys_with_delay('down', 'down', 'end', delay=0.14)
    
    # alright here I think I'd like to take the whole line instead af then take the part before the comma

    highlight_home_line()
        
    job_loc = copy_clipboard()

    job_loc = job_loc.split(',')[0]
    print(job_loc)

    
    time.sleep(1)
    pyautogui.press('f12')
    time.sleep(2)
    pyautogui.write(contractor+ ' ' + job_name + ' ' + job_loc)
    time.sleep(1)

NameSave()
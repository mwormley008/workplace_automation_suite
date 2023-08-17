import pyautogui
import time
from pyautogui import press, write, hotkey
from time import sleep

# Define a helper function to send keys with a delay
def send_keys_with_delay(*keys, delay=0.01):
    pyautogui.press(keys, interval=delay)

def repeat_hotkey(times, delay, *keys):
    for _ in range(times):
        pyautogui.hotkey(*keys)
        sleep(delay)

## Add a focus function for .doc files


# Start the flow
time.sleep(2)  # WAIT 2

# Ctrl + F
pyautogui.hotkey('ctrl', 'f')
write('dollars')
send_keys_with_delay('return', 'esc')


# Move cursor and select text
repeat_hotkey(6, 0.1, 'ctrl', 'right')

# def highlight_line():
#     press('numlock')
#     keyDown('shiftleft')
#     press('end')
#     keyUp('shiftleft')
#     sleep(1)
#     press('numlock')
#     sleep(1)

for i in range(3):
    press('numlock')
    hotkey('shift', 'ctrl', 'right')  # Select next 3 words
    press('numlock')
    sleep(.1)

time.sleep(.1)
pyautogui.hotkey('shift', 'ctrl', 'f5')
time.sleep(.1)
send_keys_with_delay('amnt')
pyautogui.hotkey('alt', 'a')
time.sleep(.1)
# press('end')
# write('refamnt ')  # Added space at the end as in original

# for i in range(6):
#     press('numlock')
#     hotkey('shift', 'ctrl', 'left')  # Select next 3 words
#     press('numlock')
#     sleep(.1)


# # Cut and search for 'dollars'
# pyautogui.hotkey('ctrl', 'x')
pyautogui.hotkey('ctrl', 'f')
sleep(.1)
write('dollars')
sleep(.1)
pyautogui.press('return')
sleep(.1)
pyautogui.press('esc')
sleep(.1)
pyautogui.press('left')
sleep(.1)
# Some more text manipulations
press('numlock')
pyautogui.hotkey('shift', 'home')
press('numlock')
# send_keys_with_delay('backspace', 'left', 0.14)  # 14ms delay
press('backspace')

# press('left')
# Final actions
pyautogui.hotkey('ctrl', 'f9')
write('REF amnt \*cardtext \*caps    ')
pyautogui.press('f9')
sleep(.1)
pyautogui.hotkey('ctrl', 'f')
sleep(.1)
write('dollars')
sleep(.1)
pyautogui.press('return')
sleep(.1)
pyautogui.press('esc')
sleep(.1)
pyautogui.press('left')
sleep(.1)
press('space')
print("Script completed.")

def NameSave():
    # Define a helper function to send keys with a delay
    def send_keys_with_delay(keys, delay=0.01):
        pyautogui.write(keys, interval=delay)

    # Search for "proposal submitted to:"
    pyautogui.hotkey('ctrl', 'f')
    send_keys_with_delay('proposal submitted to:')
    pyautogui.press('return')
    pyautogui.press('esc')
    send_keys_with_delay('downendshifthome', delay=0.15)
    pyautogui.hotkey('ctrl', 'c')
    send_keys_with_delay('upupctrlv ')

    # Search for "Work to be performed at:"
    pyautogui.hotkey('ctrl', 'f')
    send_keys_with_delay('Work to be performed at:')
    pyautogui.press('return')
    pyautogui.press('esc')
    send_keys_with_delay('downendshifthome', delay=0.14)
    pyautogui.hotkey('ctrl', 'c')
    send_keys_with_delay('upupendctrlv ')

    # Search for "Work to be performed at:" again
    pyautogui.hotkey('ctrl', 'f')
    send_keys_with_delay('Work to be performed at:')
    time.sleep(1)
    pyautogui.press('return')
    pyautogui.press('esc')
    time.sleep(2)
    send_keys_with_delay('downdowndownend', delay=0.14)
    send_keys_with_delay('ctrlleftleftleftshifthome', delay=0.1)
    pyautogui.hotkey('ctrl', 'c')
    send_keys_with_delay('upupupupendctrlvshifthome', delay=0.13)
    pyautogui.hotkey('ctrl', 'x')
    pyautogui.press('return')
    time.sleep(1)
    pyautogui.press('f12')
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

NameSave()
from pyautogui import press, write, hold, hotkey
from time import sleep


sleep(2)
hotkey('ctrl', 'h')
sleep(.25)
press ('(')

# sleep(.1)
hotkey('alt', 'a')
# sleep(.1)
press('space')
# sleep(.1)
press(')')
# sleep(.1)
hotkey('alt', 'a')
# sleep(.1)
press('space')
# sleep(.1)
press('esc')

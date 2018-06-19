import pyautogui
import keyboard
import time
import ctypes

price = '20'
date = '4/20'
#id_number = input('IdNumber:')

#pyautogui.MoveTo(300, 300)

pyautogui.doubleClick(x=800, y=300)
id_num = "1225539"
keyboard.write(id_num)
#pyautogui.typewrite("1235")
#pyautogui.press('a')
pyautogui.click(x=1200, y=300)
pyautogui.doubleClick(x=1200, y=435)
#time.sleep(3)
pyautogui.screenshot('screenshot40.png')

ctypes.windll.user32.SetCursorPos(3530,521)
x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png')
pyautogui.doubleClick(x + 50, y)
#for pos in pyautogui.locateAllOnScreen('someButton.png')
"""button7location = pyautogui.locateOnScreen('sc1.png', region=(352,958 377,1106))
button7x, button7y = pyautogui.center(button7location)
pyautogui.doubleClick(button7x, button7y)"""
#! python3
import pyautogui, sys
print('Press Ctrl-C to quit.')
try:
    while True:
        x, y = pyautogui.position()
        positionStr = 'X: ' + str(x).rjust(4) + ' Y: ' + str(y).rjust(4)
        print(positionStr, end='')
        print('\b' * len(positionStr), end='', flush=True)
except KeyboardInterrupt:
    print('\n')
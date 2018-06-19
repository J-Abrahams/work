import pyautogui
import keyboard


def search_id(pid):
    x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png')
    pyautogui.doubleClick(x + 50, y)
    keyboard.write(pid)
    keyboard.send('enter')
    keyboard.send('enter')
    

pid = input('Enter PID:')

search_id(pid)
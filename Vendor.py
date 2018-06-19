import keyboard
import pyautogui
import time
import pyperclip

def switch_site(site):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\windows_closed.png',
                                       region=(514, 245, 300, 300))
    if image == None:
        print("Close all windows")
        raise SystemExit(0)

    pyautogui.click(10, 25)
    pyautogui.click(10, 150)
    pyautogui.click(1000, 375)
    if site == 2:
        keyboard.write('A1')
        keyboard.send('enter')
    elif site == 3:
        keyboard.write('A2')
        keyboard.send('enter')
    elif site == 4:
        keyboard.write('A3')
        keyboard.send('enter')
    elif site == 5:
        keyboard.write('T')
        keyboard.send('enter')
    elif site == 8:
        keyboard.write('C')
        keyboard.send('enter')
    elif site == 9:
        keyboard.write('Welk Resort N')
        keyboard.send('enter')
    elif site == 11:
        keyboard.write('Welk Resort Bre')
        keyboard.send('enter')
    pyautogui.click(10, 25)
    pyautogui.click(10, 45)


def create_prospect()


prospect = {'site':5, 'first':'', 'last':'', 'street'}
switch_site(11)



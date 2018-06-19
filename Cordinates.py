import pyautogui
import keyboard

prospect_search = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\prospectsearch.png',
                                       region=(514, 245, 889, 566))
while prospect_search == None:
    prospect_search = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\prospectsearch.png',
                                           region=(514, 245, 889, 566))

x, y = prospect_search
pyautogui.doubleClick(x + 50, y + 42)
keyboard.write('1329146')
pyautogui.click(x + 650, y + 40)
pyautogui.click(x + 380, y + 540)

prospect_record = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changingaprospect.png',
                                       region=(514, 245, 889, 566))
while prospect_search == None:
    prospect_search = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changingaprospect.png',
                                           region=(514, 245, 889, 566))

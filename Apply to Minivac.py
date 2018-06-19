import pyautogui
import keyboard

def apply_to_mv(amount):
    x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_changing_tour_package_menu.png',
                                           region=(514, 245, 889, 566))
    pyautogui.doubleClick(x + 200, y + 220)


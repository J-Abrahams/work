import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8
import pyautogui
import keyboard


def search_pid(pid_number):
    get_m1_coordinates()
    pyautogui.doubleClick(m1['search'])
    keyboard.write(pid_number)
    pyautogui.click(m1['find_now'])
    pyautogui.click(m1['change'])


def select_tour():
    get_m2_coordinates()
    x, y = m2['title']
    # Checks if there is an audition
    audition = pyautogui.pixelMatchesColor(x + 465, y + 65, (255, 255, 255))
    while audition is True:
        y = y + 13
        audition = pyautogui.pixelMatchesColor(x + 465, y + 65, (0, 0, 0))
    pyautogui.doubleClick(x + 469, y + 67)  # Selects the top tour that isn't an audition
    # Checks if "You need to change sites" message comes up
    time.sleep(1)
    pyautogui.click(m2['yes_change_sites'])
        

def create_dep():
    x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png', region=(514, 245, 889, 566))
    pyautogui.click(x + 300, y + 20)
    try:
        x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_deposit.png', region=(514, 245, 889, 566))
    
    except Exception:
        amount = input("Deposit amount:")

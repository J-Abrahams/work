import keyboard
import pyautogui
import time
import pyperclip


def search_pid(pid):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    pyautogui.doubleClick(x + 50, y)
    keyboard.write(pid)
    pyautogui.click(x + 650, y)
    pyautogui.click(x + 400, y + 500)


def select_tour():
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    audition = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_audition.png',
                                              region=(514, 245, 889, 566))
    if audition is None:
        pyautogui.doubleClick(x, y + 12)

    else:
        pyautogui.doubleClick(x, y + 24)
    # Checks if "You need to change sites" message comes up
    time.sleep(1)
    pyautogui.click(900, 570)


def audit_refund(department):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    pyautogui.click(x + 300, y + 18)  # Tour Packages TabD
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_refundable50.png',
                                           region=(514, 245, 1000, 566))
    attempt = 0
    while image is None and attempt <= 3:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_deposit.png',
                                               region=(514, 245, 1000, 566))
        attempt += 1
    if image is None:
        print('Deposit is not $50')
    else:
        pyautogui.click(x + 375, y + 18)
        dep = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_dep50.png',
                                             region=(514, 245, 1000, 566))
        if dep is not None:
            print('DEP already there.')
            return
        dep_selected = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_dep50_selected.png',
                                                      region=(514, 245, 1000, 566))
        print(dep_selected)
        if dep_selected is not None:
            print('DEP already there.')
            return
        pyautogui.click(x + 275, y + 185)
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changingpremium.png',
                                               region=(514, 245, 889, 566))
        while image is None:
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changingpremium.png',
                region=(514, 245, 889, 566))
        x_1, y_1 = image
        pyautogui.click(x_1 + 175, y_1 + 80)
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_premium_search.png',
                                               region=(514, 245, 889, 566))
        while image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_premium_search.png',
                                                   region=(514, 245, 889, 566))
        x_2, y_2 = image
        pyautogui.click(x_2 + 80, y_2 + 40)
        keyboard.write('DEP $50')
        keyboard.send('enter')
        time.sleep(0.5)
        pyautogui.doubleClick(x_2 + 80, y_2 + 135)
        pyautogui.click(x_1 + 150, y_1 + 250)
        if department == 'ao':
            for i in range(18):
                keyboard.write('a')
        elif department == 'aj':
            for i in range(9):
                keyboard.write('a')
        elif department == 'at':
            for i in range(26):
                keyboard.write('a')
        elif department == 'ae':
            for i in range(3):
                keyboard.write('a')
        elif department == 'ag':
            for i in range(6):
                keyboard.write('a')
        elif department == 'am':
            for i in range(10):
                keyboard.write('a')
        elif department == 'cm':
            for i in range(3):
                keyboard.write('c')
        pyautogui.click(x_1 + 75, y_1 + 500)
        pyautogui.click(x + 265, y + 475)
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                               region=(514, 245, 889, 566))
        while image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                                   region=(514, 245, 889, 566))
        x, y = image
        pyautogui.click(x - 20, y + 425)


pids = [1414618, 1414679, 1414715, 1414748]
for pid in map(str, pids):
    search_pid(pid)
    select_tour()
    #  San Diego is cm
    audit_refund('ae')

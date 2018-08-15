import keyboard
import pyautogui
import time
import pyperclip
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8
import pandas as pd
import mss
import mss.tools
import datetime
import sys


def search_pid(pid_number):
    sc.get_m1_coordinates()
    pyautogui.doubleClick(m1['search'])
    keyboard.write(pid_number)
    pyautogui.click(m1['find_now'])
    pyautogui.click(m1['change'])


def create_data_frame():
    """
    Takes screenshots of the tours. Turns the screenshots into a list of dictionaries 'd'. Turns 'd' into a dataframe.
    d is a list of dictionaries such as [{''Date': '7/06/18', 'Tour_Type': 'Audition', 'Tour_Status': 'Showed'},
    {'Date': '7/06/18', 'Tour_Type': 'minivac', 'Tour_Status': 'Showed'}]
    :return:
    """
    sc.get_m2_coordinates()
    d = []
    x, y = m2['title']
    for i in range(8):
        with mss.mss() as sct:
            monitor = {'top': y + 63, 'left': x + 330, 'width': 52, 'height': 10}
            im = sct.grab(monitor)
            try:
                screenshot = sc.dates[str(mss.tools.to_png(im.rgb, im.size))]
                date = datetime.datetime.strptime(screenshot, "%m/%d/%y")
            except KeyError:
                date = None
            monitor = {'top': y + 63, 'left': x + 402, 'width': 14, 'height': 10}
            im = sct.grab(monitor)
            try:
                screenshot_2 = sc.m2_tour_types[str(mss.tools.to_png(im.rgb, im.size))]
            except KeyError:
                print(str(mss.tools.to_png(im.rgb, im.size)))
                screenshot_2 = None
            monitor = {'top': y + 63, 'left': x + 484, 'width': 14, 'height': 10}
            im = sct.grab(monitor)
            try:
                screenshot_3 = sc.m2_tour_status[str(mss.tools.to_png(im.rgb, im.size))]
            except KeyError:
                print(mss.tools.to_png(im.rgb, im.size))
                screenshot_3 = None
            y += 13
            if screenshot_2 != 'Nothing':
                try:
                    # Where the screenshots get turned into dictionaries.
                    d.append({'Date': date, 'Tour_Type': screenshot_2, 'Tour_Status': screenshot_3})
                except NameError:
                    pass
    df = pd.DataFrame(d)  # Turn d into a dataframe
    df = df[['Date', 'Tour_Type', 'Tour_Status']]  # Reorders the columns in the dataframe.
    print(df)
    return df


def select_tour(df):
    x, y = m2['title']
    tour_number = df[(df.Tour_Type != 'Audition')].index[0]
    pyautogui.doubleClick(x + 469, y + 67 + 13 * tour_number)
    # Checks if "You need to change sites" message comes up
    time.sleep(1)
    pyautogui.click(m2['yes_change_sites'])


def audit_refund(department):
    sc.get_m3_coordinates()
    pyautogui.click(m3['tour_packages'])
    time.sleep(1)
    with mss.mss() as sct:
        # The screen part to capture
        x, y = m3['title']
        monitor = {'top': y + 68, 'left': x + 464, 'width': 37, 'height': 11}
        im = sct.grab(monitor)
        try:
            amount = sc.deposit_amount[(mss.tools.to_png(im.rgb, im.size))]
        except KeyError:
            sys.exit('Deposit is not $50')
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


pids = [1425546]
for pid in map(str, pids):
    search_pid(pid)
    df = create_data_frame()
    select_tour(df)
    #  San Diego is cm
    audit_refund('ao')

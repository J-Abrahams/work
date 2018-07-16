import keyboard
import pyautogui
import time
import pyperclip
import csv
import pandas as pd
import mss
import mss.tools
from tkinter import Tk
from screenshot_data import *

m1 = {}
m2 = {}
m3 = {}


def get_m1_coordinates():
    global m1
    m1_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png',
                                              region=(514, 245, 889, 566))
    while m1_title is None:
        m1_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png',
                                                  region=(514, 245, 889, 566))
    m1['search'] = (m1_title[0] + 50, m1_title[1])
    m1['find_now'] = (m1_title[0] + 650, m1_title[1])
    m1['change'] = (m1_title[0] + 400, m1_title[1] + 500)
    return m1


def get_m2_coordinates():
    global m2
    m2_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changing_a_prospect.png',
                                              region=(514, 245, 889, 566))
    while m2_title is None:
        m2_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles'
                                                  '\\changing_a_prospect.png',
                                                  region=(514, 245, 889, 566))
    m2['title'] = (m2_title[0], m2_title[1])
    m2['first_tour'] = (m2_title[0] + 375, m2_title[1] + 65)
    m2['second_tour'] = (m2_title[0] + 375, m2_title[1] + 78)
    m2['third_tour'] = (m2_title[0] + 375, m2_title[1] + 91)
    m2['yes_change_sites'] = (m2_title[0] + 275, m2_title[1] + 266)
    m2['last_name'] = (m2_title[0] + 129, m2_title[1] + 85)
    m2['first_name'] = (m2_title[0] + 271, m2_title[1] + 85)
    m2['salutation'] = (m2_title[0] + 273, m2_title[1] + 110)
    m2['company'] = (m2_title[0] + 271, m2_title[1] + 137)
    m2['address'] = (m2_title[0] + 271, m2_title[1] + 161)
    m2['city'] = (m2_title[0] + 129, m2_title[1] + 223)
    m2['county'] = (m2_title[0] + 271, m2_title[1] + 223)
    m2['state'] = (m2_title[0] + 103, m2_title[1] + 249)
    m2['postal_code'] = (m2_title[0] + 271, m2_title[1] + 249)
    m2['country'] = (m2_title[0] + 174, m2_title[1] + 276)
    m2['phone1'] = (m2_title[0] + 129, m2_title[1] + 302)
    m2['phone2'] = (m2_title[0] + 271, m2_title[1] + 302)
    m2['fax'] = (m2_title[0] + 129, m2_title[1] + 328)
    m2['email'] = (m2_title[0] + 271, m2_title[1] + 353)
    return m2


def get_m3_coordinates():
    global m3
    m3_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\titles\\changing_a_tour.png',
                                              region=(514, 245, 889, 566))
    while m3_title is None:
        m3_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\titles\\changing_a_tour.png',
                                                  region=(514, 245, 889, 566))
    m3['title'] = (m3_title[0], m3_title[1])
    m3['tour'] = (m3_title[0] - 45, m3_title[1] + 20)
    m3['user_fields'] = (m3_title[0] + 5, m3_title[1] + 20)
    m3['notes'] = (m3_title[0] + 55, m3_title[1] + 20)
    m3['accommodations'] = (m3_title[0] + 242, m3_title[1] + 24)
    m3['tour_packages'] = (m3_title[0] + 324, m3_title[1] + 24)
    m3['premiums'] = (m3_title[0] + 397, m3_title[1] + 24)
    m3['personnel'] = (m3_title[0] + 225, m3_title[1] + 217)
    m3['prospect'] = (m3_title[0] + 143, m3_title[1] + 45)
    m3['prospect_id'] = (m3_title[0] + 84, m3_title[1] + 69)
    m3['tour_id'] = (m3_title[0] + 84, m3_title[1] + 89)
    m3['campaign'] = (m3_title[0] + 166, m3_title[1] + 122)
    m3['tour_type'] = (m3_title[0] + 143, m3_title[1] + 149)
    m3['tour_status'] = (m3_title[0] + 143, m3_title[1] + 175)
    m3['tour_date'] = (m3_title[0] + 143, m3_title[1] + 203)
    m3['tour_location'] = (m3_title[0] + 143, m3_title[1] + 227)
    m3['wave'] = (m3_title[0] + 143, m3_title[1] + 251)
    m3['team'] = (m3_title[0] + 143, m3_title[1] + 278)
    m3['insert'] = (m3_title[0] + 314, m3_title[1] + 439)
    #  Tour Packages Tab
    m3['deposit_1'] = (m3_title[0] + 563, m3_title[1] + 71)
    m3['deposit_2'] = (m3_title[0] + 563, m3_title[1] + 94)
    #  Premiums Tab
    m3['premium_1'] = (m3_title[0] + 563, m3_title[1] + 64)
    m3['premium_2'] = (m3_title[0] + 563, m3_title[1] + 77)
    m3['premium_3'] = (m3_title[0] + 563, m3_title[1] + 90)
    m3['premium_4'] = (m3_title[0] + 563, m3_title[1] + 103)
    m3['premium_5'] = (m3_title[0] + 563, m3_title[1] + 116)
    m3['premium_6'] = (m3_title[0] + 563, m3_title[1] + 129)
    return m3


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


def check_for_refundable_deposit():
    pyautogui.click(m3['tour_packages'])
    x, y = m3['deposit_1']
    is_deposit_blue = pyautogui.pixelMatchesColor(x, y, (8, 36, 107))
    while is_deposit_blue is True:
        pyautogui.tripleClick(x, y)
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changing_tour_package.png',
            region=(514, 245, 889, 566))
        while image is None:
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changing_tour_package.png',
                region=(514, 245, 889, 566))
        x_2, y_2 = image
        pyautogui.click(x_2 + 150, y_2 + 125)  # Description
        keyboard.send('ctrl + z')  # Select all
        keyboard.send('ctrl + c')  # Copy description
        r = Tk()
        result = r.selection_get(selection="CLIPBOARD")
        if 'ref' in result.lower():
            pyautogui.tripleClick(x_2 + 150, y_2 + 190)
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\view_record.png', region=(514, 245, 889, 566))
            while image is None:
                image = pyautogui.locateCenterOnScreen(
                    'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\view_record.png', region=(514, 245, 889, 566))
            x_3, y_3 = image
            pyautogui.doubleClick(x_3 + 80, y_3 + 118)
            keyboard.send('ctrl + c')
            r = Tk()
            price = str(r.selection_get(selection="CLIPBOARD")[1:3])
            pyautogui.click(x_3 + 300, y_3 + 215)
            pyautogui.click(x_2 + 350, y_2 + 400)
            return price
        else:
            pyautogui.click(x_2 + 350, y_2 + 400)
            y += 13
            pyautogui.click(x, y)
            is_deposit_blue = pyautogui.pixelMatchesColor(x, y, (8, 36, 107))

    """x, y = m3['title']
    refundable = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\refundable.png',
                                                region=(514, 245, 889, 566))
    if refundable is None:
        refundable = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\refundable.png',
            region=(514, 245, 889, 566))
    if refundable is not None:
        with mss.mss() as sct:
            monitor = {'top': y + 68, 'left': x + 461, 'width': 100, 'height': 9}
            im = sct.grab(monitor)
            price = prices[str(mss.tools.to_png(im.rgb, im.size))[106:115]]
            print(price)"""


def check_for_dep_premium():
    get_m3_coordinates()
    price = check_for_refundable_deposit()
    premiums = check_for_duplicate_premiums()
    if price == '40':
        if any(dep_40_cc in s for s in premiums) or any(dep_40_cash in s for s in premiums) or \
                any(d40_cc_dep in s for s in premiums) or any(d40_dep in s for s in premiums):
            print('\x1b[6;30;42m' + '$40 DEP is present' + '\x1b[0m')
        else:
            print('\x1b[6;30;41m' + 'Missing $40 DEP' + '\x1b[0m')
    elif price == '50':
        if any(dep_50_cc in s for s in premiums) or any(dep_50_cc in s for s in premiums):
            print('\x1b[6;30;42m' + '$50 DEP is present' + '\x1b[0m')
        else:
            print('\x1b[6;30;41m' + 'Missing $50 DEP' + '\x1b[0m')


def check_for_duplicate_premiums():
    premiums = []
    pyautogui.click(m3['premiums'])
    time.sleep(0.3)
    pyautogui.click(m3['premium_1'])
    x, y = m3['premium_1']
    number_of_premiums = 0
    is_premium_blue = pyautogui.pixelMatchesColor(x, y, (8, 36, 107))
    while is_premium_blue is True:
        number_of_premiums += 1
        with mss.mss() as sct:
            monitor = {'top': y - 4, 'left': x - 223, 'width': 100, 'height': 9}
            im = sct.grab(monitor)
            premiums.append(str(mss.tools.to_png(im.rgb, im.size)))
            print((mss.tools.to_png(im.rgb, im.size)))
        y += 13
        pyautogui.click(x, y)
        time.sleep(0.3)
        is_premium_blue = pyautogui.pixelMatchesColor(x, y, (8, 36, 107))
        if is_premium_blue is False:
            is_premium_blue = pyautogui.pixelMatchesColor(x, y, (8, 36, 107))
    if len(premiums) != len(set(premiums)):
        print('\x1b[6;30;41m' + str(number_of_premiums) + ' Premiums - DUPLICATES' + '\x1b[0m')
    else:
        print('\x1b[6;30;42m' + str(number_of_premiums) + ' Premiums - No Duplicates' + '\x1b[0m')
    return premiums


def confirm_tour_status(status):
    x, y = m3['title']
    if status == 'c':
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\confirmed.png',
                                               region=(x + 27, y + 132, 131, 103))
        attempts = 0
        while image is None and attempts <= 2:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\confirmed.png',
                                                   region=(514, 245, 889, 566))
            attempts += 1
        if image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\showed.png',
                                                   region=(x + 27, y + 132, 131, 103))
        if image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\on_tour.png',
                                                   region=(x + 27, y + 132, 131, 103))
        if image is None:
            print('\x1b[6;30;41m' + 'TOUR STATUS MIGHT BE INCORRECT' + '\x1b[0m')
        else:
            print('\x1b[6;30;42m' + 'Tour status is good' + '\x1b[0m')
    elif status == 'r':
        if pyautogui.pixelMatchesColor(x + 48, y + 171, (0, 0, 0)) is True:
            print('\x1b[6;30;42m' + 'Tour status is good' + '\x1b[0m')
        elif pyautogui.pixelMatchesColor(x + 99, y + 171, (0, 0, 0)) is True:
            print('\x1b[6;30;42m' + 'Tour status is good' + '\x1b[0m')
        else:
            print('\x1b[6;30;41m' + 'TOUR STATUS MIGHT BE INCORRECT' + '\x1b[0m')
    elif status == 'x':
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\canceled.png',
                                               region=(x + 27, y + 132, 131, 103))
        attempts = 0
        while image is None and attempts <= 2:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\canceled.png',
                                                   region=(x + 27, y + 132, 131, 103))
            attempts += 1
        if image is None:
            print('\x1b[6;30;41m' + 'TOUR STATUS MIGHT BE INCORRECT' + '\x1b[0m')
        else:
            print('\x1b[6;30;42m' + 'Tour status is good' + '\x1b[0m')


def confirm_sol_in_userfields(sol):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    pyautogui.click(x, y + 18)  # User Fields Tab
    try:
        x_1, y_1 = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_confirmer_sol.png',
                                                  region=(514, 245, 889, 566))
        pyautogui.doubleClick(x + 115, y + 222)
        keyboard.press_and_release('ctrl + c')
        tsw_sol = str(pyperclip.paste())
        print(tsw_sol + " changed to " + sol)
        pyperclip.copy(sol)
        keyboard.press_and_release('ctrl + v')

    except TypeError:
        print('\x1b[6;30;42m' + 'Sol number is good' + '\x1b[0m')

    pyautogui.click(x - 65, y + 18)


def check_deposit():
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    pyautogui.click(x + 300, y + 18)


def enter_personnel(sol, status):
    get_m3_coordinates()
    if status == 'c':
        confirm_sol_in_userfields(sol)
    pyautogui.click(m3['personnel'])
    pyautogui.click(m3['insert'])
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                               region=(514, 245, 889, 566))
    x_1, y_1 = image
    time.sleep(0.3)
    pyautogui.click(x_1 + 75, y_1 + 25)  # By Personnel Number Tab
    keyboard.write(sol)
    pyautogui.doubleClick(x_1, y_1 + 100)  # Person in list
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_personnel_titles_menu.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_personnel_titles_menu.png',
                                               region=(514, 245, 889, 566))
    x_3, y_3 = image
    pyautogui.click(x_3 + 75, y_3 + 150)  # Close
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_addingrecord.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_addingrecord.png',
                                               region=(514, 245, 889, 566))
    x_4, y_4 = image
    try:
        x_5, y_5 = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_confirmer.png',
                                                  region=(514, 245, 889, 566))
    except TypeError:
        pyautogui.click(x_4 + 90, y_4 + 80)
        keyboard.write("cc")
    pyautogui.click(x_4 + 90, y_4 + 105)
    if status == 'c':
        keyboard.write("cc")
    elif status == 'r':
        keyboard.write("r")
    elif status == 'x':
        keyboard.write("c")
    elif status == 'u':
        keyboard.write("u")
    elif status == 'tav':
        keyboard.write("t")
    pyautogui.click(x_4 + 90, y_4 + 350)


def convert_excel_to_csv():
    xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\Downloads\\1.xlsx")
    df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
    df.to_csv('file.csv')


def activation_sheet():
    for pid in pids:
        if pid != '':
            search_pid(pid)
            select_tour()
            check_for_duplicate_premiums()
            keep_going = input("Everything ok?")
            if keep_going != '':
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                       region=(514, 245, 889, 566))
                while image is None:
                    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                           region=(514, 245, 889, 566))
                x, y = image
                pyautogui.click(x + 265, y + 475)
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                                       region=(514, 245, 889, 566))
                while image is None:
                    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                                           region=(514, 245, 889, 566))
                x, y = image
                pyautogui.click(x - 20, y + 425)


def manual_confirmation(pids):
    for pid in pids:
        if pid != '':
            search_pid(pid)
            select_tour()
            check_for_dep_premium()
            input("Everything ok?")
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                   region=(514, 245, 889, 566))
            while image is None:
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                       region=(514, 245, 889, 566))
            x, y = image
            pyautogui.click(x + 265, y + 475)
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                                   region=(514, 245, 889, 566))
            while image is None:
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                                       region=(514, 245, 889, 566))
            x, y = image
            pyautogui.click(x - 20, y + 425)


def automatic_confirmation():
    convert_excel_to_csv()
    with open('file.csv') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            pids = row['PID'].replace('.0', '')
            search_pid(pids)
            select_tour()
            check_for_dep_premium()
            input("Everything ok?")
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                   region=(514, 245, 889, 566))
            while image is None:
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                       region=(514, 245, 889, 566))
            x, y = image
            pyautogui.click(x + 265, y + 475)
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                                   region=(514, 245, 889, 566))
            while image is None:
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                                       region=(514, 245, 889, 566))
            x, y = image
            pyautogui.click(x - 20, y + 425)


pids = ['1403818', '625554', '1220807', '35897', '1255113', '658248', '507196', '1309578', '1357445', '705065',
        '653771', '1417501', '620362', '1412509', '1417790', '1415366', '1411611', '1415513', '1415127', '1417426',
        '1406230', '1413582', '1419505', '1418336', '1401205', '1417952']

auto_or_manual = input('Auto (A) or Manual (M):')
print("Justin Locke's SOL is SOL4967")
sol = "SOL" + input("SOL #:")
if auto_or_manual == 'a' or auto_or_manual == 'A':
    automatic_confirmation()
else:
    manual_confirmation(pids)

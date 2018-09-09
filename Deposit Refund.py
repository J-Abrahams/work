import csv
import time
from tkinter import Tk
import keyboard
import mss
import mss.tools
import pandas as pd
import pyautogui
import clipboard
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10
import datetime
from tabulate import tabulate
import sys
import pickle
import openpyxl
import re
import random

# import importlib
# importlib.reload(sc)
transaction_code = 0


# TODO Make the program handle prev tours automatically.
# TODO get the reference number. 1312792
# TODO 1224070
# TODO Fix data frame on PID 1425576


def take_screenshot(x, y, width, height, save_file=False):
    with mss.mss() as sct:
        monitor = {'top': y, 'left': x, 'width': width, 'height': height}
        im = sct.grab(monitor)
        screenshot = str(mss.tools.to_png(im.rgb, im.size))
        if save_file:
            now = datetime.datetime.now()
            output = now.strftime("%m-%d-%H-%M-%S.png".format(**monitor))
            mss.tools.to_png(im.rgb, im.size, output=output)
        return screenshot


def read_pickle_file(file_name):
    with open('text_files\\' + file_name, 'rb') as file:
        return pickle.load(file)


def get_date(month, day, year):
    f = open('text_files\\dates.p', 'rb')
    date_dictionary = pickle.load(f)
    f.close()
    try:
        month = date_dictionary[month]
        day = date_dictionary[day]
        year = date_dictionary[year]
        if month not in ['Nothing', 'Error']:
            tour_date = ('{}/{}/{}'.format(month, day, year))
            return tour_date
            # datetime.datetime.strptime(tour_date, "%m/%d/%Y")
        elif month == 'Error':
            return 'Error'
        elif month == 'Nothing':
            return 'Nothing'
    except KeyError:
        return 'Nothing'


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
    pretty_d = []
    x, y = m2['title']
    for i in range(8):
        month = take_screenshot(x + 327, y + 63, 13, 10)
        day = take_screenshot(x + 342, y + 63, 15, 10)
        year = take_screenshot(x + 358, y + 63, 27, 10)
        tour_date = get_date(month, day, year)
        tour_type = take_screenshot(x + 402, y + 63, 14, 10)
        tour_status = take_screenshot(x + 484, y + 63, 14, 10)
        try:
            tour_type = sc.m2_tour_types[tour_type]
        except KeyError:
            print('Unrecognized tour type')
            print(tour_type)
            tour_type = None
        try:
            tour_status_dict = read_pickle_file('m2_tour_status.p')
            tour_status = tour_status_dict[tour_status]
        except KeyError:
            print('Unrecognized tour status')
            print(tour_status)
            tour_status = None
        y += 13
        if tour_date != 'Nothing':
            try:
                # Where the screenshots get turned into dictionaries.
                pretty_d.append({'Date': tour_date, 'Tour_Type': tour_type, 'Tour_Status': tour_status})
                tour_date = pd.to_datetime(tour_date)
                d.append({'Date': tour_date, 'Tour_Type': tour_type, 'Tour_Status': tour_status})
            except NameError:
                pass
    df = pd.DataFrame(d)  # Turn d into a dataframe
    df = df[['Date', 'Tour_Type', 'Tour_Status']]  # Reorders the columns in the dataframe.
    pretty_df = pd.DataFrame(pretty_d)
    pretty_df = pretty_df[['Date', 'Tour_Type', 'Tour_Status']]
    print(tabulate(pretty_df, headers='keys', tablefmt='psql'))
    return df


def select_tour(df, attempt_number, date):
    x, y = m2['title']
    date = date + '/18'
    date = datetime.datetime.strptime(date, "%m/%d/%y")
    # Returns the top tour that is Showed, not an Audition, and at most a week before the date we entered.
    # tour_number is the index of the correct tour. Ex: 1 if the second tour is the correct one.
    try:
        tour_number = df[(df.Tour_Status == 'Showed') & (df.Tour_Type != 'Audition') &
                         ((date - df.Date) <= datetime.timedelta(days=7))].index[attempt_number - 1]
    except IndexError:
        tour_number = df[(df.Tour_Status == 'No_Show') & (df.Tour_Type != 'Audition')].index[0]
    pyautogui.doubleClick(x + 469, y + 67 + 13 * tour_number)
    # Checks if "You need to change sites" message comes up
    time.sleep(1)
    pyautogui.click(m2['yes_change_sites'])


def double_check_pid(pid_number):
    sc.get_m2_coordinates()
    pyautogui.doubleClick(m2['prospect_id'])
    keyboard.send('ctrl + c')
    copied_text = clipboard.paste()
    for i in range(3):
        if copied_text != pid_number:
            time.sleep(0.3)
            pyautogui.doubleClick(m2['prospect_id'])
            keyboard.send('ctrl + c')
            copied_text = clipboard.paste()
    if copied_text != pid_number:
        input('Is the pid correct?')
        return
    if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\company.png',
                                      region=(514, 245, 889, 566)) is not None:
        return
    pyautogui.click(m2['company'])
    keyboard.send('ctrl + z')
    time.sleep(1)
    keyboard.send('ctrl + c')
    copied_text = clipboard.paste()
    if 'pid' in copied_text.lower():
        input('Is the pid correct?')
        return


def count_deposit_items():
    sc.get_m6_coordinates()
    number_of_deposit_items = 0
    x, y = m6['title']
    while True:
        with mss.mss() as sct:
            monitor = {'top': y + 189, 'left': x + 339, 'width': 10, 'height': 8}
            im = sct.grab(monitor)
            screenshot = str(mss.tools.to_png(im.rgb, im.size))
        if screenshot == sc.no_deposit_items:
            return number_of_deposit_items
        else:
            number_of_deposit_items += 1
            y += 13


def change_deposit_title(price, cash=None):
    """
    Checks to make sure that the deposit price is correct. Ex: If the sheet says $50, then this makes sure that the
    deposit is also $50.
    Changes the title from 'Refundable' to 'Refunded'
    :rtype: 'prev', 'ams', 'ir', 'sol', 'ih'
    """
    sc.get_m3_coordinates()
    amount = 0
    old_title = 'old'
    x, y = m3['title']
    x_2, y_2 = m3['deposit_1']
    attempts = 0
    while amount != price and attempts <= 2:
        pyautogui.click(m3['tour_packages'])
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\balance.png',
                                               region=(700, 245, 850, 566))
        while image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\balance.png',
                                                   region=(700, 245, 850, 566))
        pyautogui.click(x_2, y_2)
        pyautogui.click(m3['change_deposit'])
        deposit_item_amount = count_deposit_items()
        sc.get_m6_coordinates()
        x, y = m6['deposit_1']
        y = y + 13 * (deposit_item_amount - 1)
        pyautogui.click(x, y)
        time.sleep(0.3)
        with mss.mss() as sct:
            # Takes screenshot of lowest amount inside of the deposit
            x, y = m6['title']
            y = y + 13 * (deposit_item_amount - 1)
            monitor = {'top': y + 189, 'left': x + 185, 'width': 33, 'height': 8}
            im = sct.grab(monitor)
            try:
                amount = sc.deposit_item_amount[str(mss.tools.to_png(im.rgb, im.size))]
            except KeyError:
                amount = 0
                print('Don\'t recognize the amount')
                print(mss.tools.to_png(im.rgb, im.size))
                output = 'monitor-1-crop.png'
                mss.tools.to_png(im.rgb, im.size, output=output)
        attempts += 1
        if amount != price:
            pyautogui.click(m6['ok'])
            y_2 += 13
    if amount != price:
        pyautogui.click(m3['ok'])
        return 'fail'
    attempts = 0
    while 'ref' not in old_title.lower() and attempts <= 3:
        pyautogui.click(m6['description'])
        keyboard.send('ctrl + z')
        keyboard.send('ctrl + c')
        old_title = clipboard.paste()
        attempts += 1
        print(old_title.lower())
    if 'ref' not in old_title.lower():
        if old_title.lower() == 'ams dep':
            old_title = 'AMS/Refunded Deposit'
        else:
            sys.exit("Wrong Title")
    new_title = old_title.replace("able", "ed")
    new_title = new_title.replace("ABLE", "ED")
    new_title = new_title.replace(" /", "/")
    new_title = new_title.replace("/ ", "/")
    keyboard.write(new_title)
    if 'prev' in new_title.lower() and cash is None:
        return 'prev'
    elif 'ir' in new_title.lower() and 'refunded' in new_title.lower():
        return "ir"
    elif 'ams' in new_title.lower() and 'refunded' in new_title.lower():
        return "ams"
    elif 'sol' in new_title.lower() and 'refunded' in new_title.lower():
        return "sol"
    elif 'ih' in new_title.lower() and 'refunded' in new_title.lower():
        return 'ih'


def copy_reference_number():
    deposit_item_amount = count_deposit_items()
    sc.get_m6_coordinates()
    x, y = m6['deposit_1']
    y = y + 13 * (deposit_item_amount - 1)
    pyautogui.click(x, y)
    keyboard.send('alt + v')
    sc.get_m7_coordinates()
    pyautogui.doubleClick(m7['reference'])
    keyboard.press_and_release('ctrl + c')
    time.sleep(0.5)
    old_reference = clipboard.paste()
    old_reference = old_reference.upper()
    if old_reference[0] != 'D':
        sys.exit("Wrong Reference")
    new_reference = old_reference.replace("D-", "R-")
    clipboard.copy(str(new_reference))
    pyautogui.click(m7['cancel'])


def ams_credit_refund(date):
    pyautogui.click(m6['payment'])
    pyautogui.click(m8['transaction_code'])
    attempts = 0
    image = pyautogui.locateCenterOnScreen(
        'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(136, 652, 392, 247))
    while image is None and attempts <= 2:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(136, 652, 392, 247))
        attempts += 1
    if image is None:
        ams_credit_refund = 0
        return ams_credit_refund
    else:
        pyautogui.click(image)
        ams_credit_refund = 1
        return ams_credit_refund


def select_ams_refund_payment(date, price, description, reference_number=None):
    sc.get_m6_coordinates()
    attempts = 0
    change_description_name = 0
    image = None
    global transaction_code
    if description == 'ams':
        if transaction_code > 4:
            transaction_code = 0
        if transaction_code == 0 or transaction_code == 1:
            pyautogui.click(m6['payment'])
            sc.get_m8_coordinates()
            pyautogui.click(m8['transaction_code'])
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(136, 652, 392, 247))
            while image is None and attempts <= 2:
                image = pyautogui.locateCenterOnScreen(
                    'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(136, 652, 392, 247))
                attempts += 1
            if image is not None:
                transaction_code = 1
            else:
                transaction_code = 0
                pyautogui.click(m8['cancel'])
        if transaction_code == 0 or transaction_code == 2:
            attempts = 0
            pyautogui.click(m6['payment'])
            sc.get_m8_coordinates()
            pyautogui.click(m8['transaction_code'])
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_cc_refund.png', region=(136, 652, 392, 247))
            while image is None and attempts <= 2:
                image = pyautogui.locateCenterOnScreen(
                    'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_cc_refund.png', region=(136, 652, 392, 247))
                attempts += 1
            if image is not None:
                transaction_code = 2
            else:
                transaction_code = 0
                pyautogui.click(m8['cancel'])
        if transaction_code == 0 or transaction_code == 3:
            attempts = 0
            pyautogui.click(m6['insert'])
            sc.get_m8_coordinates()
            pyautogui.click(m8['transaction_code'])
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_cc_ref.png', region=(136, 652, 392, 247))
            while image is None and attempts <= 2:
                image = pyautogui.locateCenterOnScreen(
                    'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_cc_ref.png', region=(136, 652, 392, 247))
                attempts += 1
            if image is not None:
                transaction_code = 3
            else:
                transaction_code = 0
                pyautogui.click(m8['cancel'])
        if transaction_code == 0 or transaction_code == 4:
            pyautogui.click(m6['insert'])
            sc.get_m8_coordinates()
            pyautogui.click(m8['transaction_code'])
            attempts = 0
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(136, 652, 392, 247))
            while image is None and attempts <= 2:
                image = pyautogui.locateCenterOnScreen(
                    'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(136, 652, 392, 247))
            if image is not None:
                transaction_code = 4
            else:
                transaction_code = 0
                sys.exit("Couldn't find correct choice")
    elif description == 'ir':
        if (0 < transaction_code < 5) or transaction_code > 6:
            transaction_code = 0
        if transaction_code == 0 or transaction_code == 5:
            attempts = 0
            pyautogui.click(m6['insert'])
            sc.get_m8_coordinates()
            pyautogui.click(m8['transaction_code'])
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\ir_refund.png', region=(136, 652, 392, 247))
            while image is None and attempts <= 2:
                image = pyautogui.locateCenterOnScreen(
                    'C:\\Users\\Jared.Abrahams\\Screenshots\\ir_refund.png', region=(136, 652, 392, 247))
                attempts += 1
            if image is not None:
                transaction_code = 5
            else:
                transaction_code = 0
                pyautogui.click(m8['cancel'])
        if transaction_code == 0 or transaction_code == 6:
            attempts = 0
            pyautogui.click(m6['payment'])
            sc.get_m8_coordinates()
            pyautogui.click(m8['transaction_code'])
            pyautogui.click(m8['transaction_code_scroll_bar'])
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\ir_cc_refund.png', region=(136, 652, 392, 247))
            while image is None and attempts <= 2:
                image = pyautogui.locateCenterOnScreen(
                    'C:\\Users\\Jared.Abrahams\\Screenshots\\ir_cc_refund.png', region=(136, 652, 392, 247))
                attempts += 1
            if image is not None:
                transaction_code = 6
            else:
                # Selecting IH Credit Refund and changing it to IR Credit Refund
                transaction_code = 0
                pyautogui.click(m8['cancel'])
                attempts = 0
                pyautogui.click(m6['insert'])
                sc.get_m8_coordinates()
                pyautogui.click(m8['transaction_code'])
                image = pyautogui.locateCenterOnScreen(
                    'C:\\Users\\Jared.Abrahams\\Screenshots\\ih_credit_refund.png', region=(136, 652, 392, 247))
                while image is None and attempts <= 2:
                    image = pyautogui.locateCenterOnScreen(
                        'C:\\Users\\Jared.Abrahams\\Screenshots\\ih_credit_refund.png', region=(136, 652, 392, 247))
                    attempts += 1
                if image is not None:
                    change_description_name = 1
                else:
                    transaction_code = 0
                    sys.exit("Couldn't find correct choice")
    elif description == 'sol':
        if 0 < transaction_code < 7:
            transaction_code = 0
        if transaction_code == 0 or transaction_code == 7:
            attempts = 0
            pyautogui.click(m6['insert'])
            sc.get_m8_coordinates()
            pyautogui.click(m8['transaction_code'])
            pyautogui.click(m8['transaction_code_scroll_bar'])
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\sol_cc_refund.png', region=(136, 652, 392, 247))
            while image is None and attempts <= 2:
                image = pyautogui.locateCenterOnScreen(
                    'C:\\Users\\Jared.Abrahams\\Screenshots\\sol_cc_refund.png', region=(136, 652, 392, 247))
                attempts += 1
            if image is not None:
                transaction_code = 7
            else:
                transaction_code = 0
                pyautogui.click(m8['cancel'])
        if transaction_code == 0 or transaction_code == 8:
            attempts = 0
            pyautogui.click(m6['insert'])
            sc.get_m8_coordinates()
            pyautogui.click(m8['transaction_code'])
            pyautogui.click(m8['transaction_code_scroll_bar'])
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\sol_credit_refund.png', region=(136, 652, 392, 247))
            while image is None and attempts <= 2:
                image = pyautogui.locateCenterOnScreen(
                    'C:\\Users\\Jared.Abrahams\\Screenshots\\sol_credit_refund.png', region=(136, 652, 392, 247))
                attempts += 1
            if image is not None:
                transaction_code = 8
            else:
                transaction_code = 0
                pyautogui.click(m8['cancel'])
        if transaction_code == 0 or transaction_code == 9:
            attempts = 0
            pyautogui.click(m6['payment'])
            sc.get_m8_coordinates()
            pyautogui.click(m8['transaction_code'])
            pyautogui.click(m8['transaction_code_scroll_bar'])
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\sol_credit_refund.png', region=(136, 652, 392, 247))
            while image is None and attempts <= 2:
                image = pyautogui.locateCenterOnScreen(
                    'C:\\Users\\Jared.Abrahams\\Screenshots\\sol_credit_refund.png', region=(136, 652, 392, 247))
                attempts += 1
            if image is not None:
                transaction_code = 9
            else:
                transaction_code = 0
                sys.exit("Couldn't find correct choice")
    pyautogui.click(image)
    if change_description_name == 1:
        pyautogui.click(m8['description'])
        keyboard.send('ctrl + z')
        keyboard.write('IR CREDIT CARD REFUND')
        pyautogui.doubleClick(m8['amount'])
        keyboard.write(price)
    if (2 < transaction_code < 6) or (6 < transaction_code < 9):
        pyautogui.doubleClick(m8['amount'])
        keyboard.write(price)
    pyautogui.doubleClick(m8['reference'])
    if reference_number is None:
        keyboard.send('ctrl + v')
    else:
        keyboard.write(reference_number)
    pyautogui.doubleClick(m8['date'])
    keyboard.write(date)
    keyboard.send('tab')
    pyautogui.click(m8['ok'])
    time.sleep(0.3)
    # time.sleep(0.3)
    pyautogui.click(880, 565)  # Clicking yes to the warning that appears
    pyautogui.click(m6['ok'])
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


def select_ir_refund_payment(date, price, reference_number=None):
    sc.get_m6_coordinates()
    image = None
    pyautogui.click(m6['insert'])
    sc.get_m8_coordinates()
    pyautogui.click(m8['transaction_code'])
    attempts = 0
    global transaction_code
    if transaction_code == 0 or transaction_code == 1:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\ir_refund.png', region=(136, 652, 392, 247))
        while image is None and attempts <= 2:
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\ir_refund.png', region=(136, 652, 392, 247))
            attempts += 1
        if image is not None:
            transaction_code = 1
        else:
            transaction_code = 0
    if transaction_code == 0 or transaction_code == 2:
        attempts = 0
        pyautogui.click(m8['cancel'])
        pyautogui.click(m6['payment'])
        sc.get_m8_coordinates()
        pyautogui.click(m8['transaction_code'])
        pyautogui.click(m8['transaction_code_scroll_bar'])
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\ir_cc_refund.png', region=(136, 652, 392, 247))
        while image is None and attempts <= 2:
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\ir_cc_refund.png', region=(136, 652, 392, 247))
            attempts += 1
        if image is not None:
            transaction_code = 2
        else:
            transaction_code = 0
            sys.exit("Couldn't find correct choice")
    pyautogui.click(image)
    if transaction_code == 1:
        pyautogui.doubleClick(m8['amount'])
        keyboard.write(price)
    pyautogui.doubleClick(m8['reference'])
    if reference_number is None:
        keyboard.send('ctrl + v')
    else:
        keyboard.write(reference_number)
    pyautogui.doubleClick(m8['date'])
    keyboard.write(date)
    keyboard.send('tab')
    pyautogui.click(m8['ok'])
    time.sleep(0.3)
    pyautogui.click(880, 565)  # Clicking yes to the warning that appears
    now = datetime.datetime.now()
    now_str = now.strftime("%Y-%m-%d-%H-%M-%S")
    outfile = pyautogui.screenshot(
        'C:\\Users\\Jared.Abrahams\\work\\deposit_screenshots\\ImageFile{}.png'.format(now_str))
    pyautogui.click(m6['ok'])
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


def convert_excel_to_csv():
    xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\Downloads\\deposit_pids.xlsx")
    df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
    df.to_csv('file.csv')


def count_number_of_pids():
    number_of_pids = 0
    with open('file.csv') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            number_of_pids += 1
    return number_of_pids


def mark_row_as_completed(index):
    wb = openpyxl.load_workbook(filename='C:\\Users\\Jared.Abrahams\\Downloads\\deposit_pids.xlsx')
    ws = wb.worksheets[0]
    ws.cell(row=int(index) + 2, column=5).value = 'x'
    wb.save('C:\\Users\\Jared.Abrahams\\Downloads\\deposit_pids.xlsx')


def show_progress(pid, progress, number_of_pids):
    percentage = round(progress * 100 / number_of_pids, 2)
    print('{} / {} - {} - {}'.format(str(progress), str(number_of_pids), str(percentage) + '%', str(pid)))


def use_excel_sheet():
    convert_excel_to_csv()
    number_of_pids = count_number_of_pids()
    progress = 1
    with open('file.csv') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            if row['Completed'] == 'x':
                progress += 1
                continue
            pids = row['PID'].replace('.0', '')
            price = row['price']
            date = row['date']
            date = datetime.datetime.strptime(date, '%Y-%m-%d').strftime('%m/%d')
            cash = row['cash']
            index = row['']
            show_progress(pids, progress, number_of_pids)
            if cash == 'x' or cash == 'X':
                cash_or_cc = 'cash'
            else:
                cash_or_cc = 'cc'
            search_pid(pids)
            double_check_pid(pids)
            df = create_data_frame()
            select_tour(df, 1, date)
            if cash_or_cc == 'cc':
                deposit_type = change_deposit_title(price)
            else:
                deposit_type = change_deposit_title(price, cash)
            if deposit_type == 'fail':
                select_tour(df, 2, date)
                if cash_or_cc == 'cc':
                    deposit_type = change_deposit_title(price)
                else:
                    deposit_type = change_deposit_title(price, cash)
            if deposit_type != 'prev' and cash_or_cc == 'cc':
                copy_reference_number()
            elif cash_or_cc == 'cash':
                clipboard.copy('R-CASH')
            else:
                with open('Deposit_Errors.txt', 'a') as out:
                    out.write('{}\n'.format(pids))
                    sc.get_m6_coordinates()
                    pyautogui.click(m6['ok'])
                    pyautogui.click(m6['ok'])
                    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                           region=(514, 245, 889, 566))
                    while image is None:
                        image = pyautogui.locateCenterOnScreen(
                            'C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                            region=(514, 245, 889, 566))
                    x, y = image
                    pyautogui.click(x + 265, y + 475)
                    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                                           region=(514, 245, 889, 566))
                    while image is None:
                        image = pyautogui.locateCenterOnScreen(
                            'C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                            region=(514, 245, 889, 566))
                    x, y = image
                    pyautogui.click(x - 20, y + 425)
                # ams_ir_or_sol = input('ams, ir, or sol:')
                # old_reference = input('Reference Number starting with D:')
                # reference_number = old_reference.replace("D-", "R-")
                # if ams_ir_or_sol == 'ams':
                #     select_ams_refund_payment(date, price, 'ams', reference_number)
                # elif ams_ir_or_sol == 'ir':
                #     select_ams_refund_payment(date, price, 'ir', reference_number)
                # elif ams_ir_or_sol == 'sol':
                #     select_ams_refund_payment(date, price, 'sol', reference_number)
            if deposit_type == 'ams':
                select_ams_refund_payment(date, price, 'ams')
            elif deposit_type == 'ir':
                select_ams_refund_payment(date, price, 'ir')
            elif deposit_type == 'sol':
                select_ams_refund_payment(date, price, 'sol')
            mark_row_as_completed(index)
            progress += 1


use_excel_sheet()

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
import core_functions as cf
import pytest
import logging

# import importlib
# importlib.reload(cf)
errors = 0
progress = 0
sol_numbers = {'Jennifer Gordon': 'SOL2956', 'Katherine England': 'SOL23521', 'Katherine Albini': 'SOL23521',
               'Katherine England/Abini': 'SOL23521', 'Justin Locke': 'SOL4967', 'Brian Bennett': 'SOL3055',
               'Carter Roedell': 'SOL23345', 'Fernanda Hernandez': 'SOL26788', 'Fern Hernandez': 'SOL26788',
               'Alton Major': 'SOL4809', 'Thuy Pham': 'SOL25688', 'Julianne Martinez': 'SOL22766',
               'Quenton Stroud': 'SOL27228', 'Sadie Oliver': 'SOL26834', 'Valeria Rebollar': 'SOL24218',
               'Sergio Espinoza': 'SOL23542', 'Olivia Larimer': 'SOL5463', 'Grayson Corbin': 'SOL1604',
               'Deonte Keller': 'SOL27498', 'Rayven Alexander': 'SOL24125', 'Deeandra Castillo': 'SOL5495',
               'Kenan Williams': 'SOL27567', 'Jenniffer Abbott': 'SOL5456', 'Met Austin Simon': 'SOL27647'}
f = open('text_files\\premiums.p', 'rb')
premium_dict = pickle.load(f)
f.close()


#  TODO Check for accommodation cancel notes 1423766
#  TODO Partial refunds 1423766
#  TODO Check title of note when it's a cancel
#  TODO Fix duplicate deposits.
#  TODO If a deposit is refunded, don't create a missing dep message 1384377
#  TODO If an accommodation is canceled, check if the cancel box is ticked 1218776
#  TODO Make program recognize an Additional Nights deposit 984256
#  TODO If there are 2 refundable deposits and 2 dep premiums, don't count the 2 dep premiums as duplicates 1408810
#  TODO Fix data frame on PID 1425576
#  TODO If there are a lot of slashes in a note, count that as a confirm note.
#  TODO Automatically fill in the tour result when it's a cancel.
#  TODO If tour is InHouse and labeled as Minivac, make it print Minivac - 0 in green because there shouldn't be any
#  TODO accommodations. 1427980
#  TODO If a person offers an upgrade and the person takes the offer within 72 hours, then the upgrade goes to the
#  TODO person who offered it. 1433656


def gather_m3_data():
    sc.get_m3_coordinates()
    x, y = m3['title']
    tour_types_dict = cf.read_pickle_file('m3_tour_type.p')
    m3_tour_type = tour_types_dict[cf.take_screenshot(x + 36, y + 143, 89, 12)]
    m3_tour_status = sc.m3_tour_status[cf.take_screenshot(x + 37, y + 170, 94, 11)]
    month = cf.take_screenshot(x + 37, y + 196, 13, 10)
    day = cf.take_screenshot(x + 52, y + 196, 15, 10)
    year = cf.take_screenshot(x + 68, y + 196, 27, 10)
    m3_date = cf.turn_screenshots_into_date(month, day, year)
    try:
        m3_date = datetime.datetime.strptime(m3_date, "%m/%d/%Y")
    except ValueError:
        month = cf.take_screenshot(x + 40, y + 196, 13, 10)
        day = cf.take_screenshot(x + 55, y + 196, 15, 10)
        year = cf.take_screenshot(x + 71, y + 196, 27, 10)
        m3_date = cf.turn_screenshots_into_date(month, day, year)
        m3_date = datetime.datetime.strptime(m3_date, "%m/%d/%Y")
    return m3_tour_type, m3_tour_status, m3_date


# m2 functions

def select_tour(status, attempt_number=1):
    x, y = m2['title']
    current_date = cf.get_current_date()
    df, pretty_df = cf.create_accommodations_dataframe()
    print(tabulate(pretty_df, headers='keys', tablefmt='psql'))
    # Returns the top tour that is Showed, not an Audition, and at most a week before the date we entered.
    # tour_number is the index of the correct tour. Ex: 1 if the second tour is the correct one.
    if 'c' in status:
        try:
            tour_number = df[((df.Tour_Status == 'Showed') | (df.Tour_Status == 'Confirmed') |
                              (df.Tour_Status == 'No_Show') | (df.Tour_Status == 'On_Tour')) &
                             (df.Tour_Type != 'Audition') &
                             ((df.Date - current_date) >= datetime.timedelta(days=-1)) &
                             ((df.Date - current_date) <= datetime.timedelta(days=14))].index[attempt_number - 1]
        except IndexError:
            print('Couldn\'t find correct tour')
            tour_number = df[(df.Tour_Type != 'Audition')].index[attempt_number - 1]
    elif 'r' in status:
        try:
            tour_number = df[((df.Tour_Status == 'Rescheduled') &
                              ((df.Date - current_date) >= datetime.timedelta(days=0))) |
                             ((df.Tour_Type == 'Open_Reservation') &
                              (df.Date == datetime.datetime.strptime('1/1/1900', "%m/%d/%Y"))) &
                             (df.Tour_Type != 'Audition')].index[attempt_number - 1]
        except IndexError:
            print('Couldn\'t find correct tour')
            tour_number = df[(df.Tour_Type != 'Audition')].index[attempt_number - 1]
    elif 'x' in status:
        try:
            tour_number = df[(df.Tour_Status == 'Canceled') & (df.Tour_Type != 'Audition')].index[attempt_number - 1]
        except IndexError:
            print('Couldn\'t find correct tour')
            tour_number = df[(df.Tour_Type != 'Audition')].index[attempt_number - 1]
    elif 'u' in status:
        try:
            tour_number = df[(df.Tour_Type == 'Minivac') &
                             ((df.Date - current_date) >= datetime.timedelta(days=-1))].index[attempt_number - 1]
        except IndexError:
            print('Couldn\'t find correct tour')
            tour_number = df[(df.Tour_Type != 'Audition')].index[attempt_number - 1]
    elif 't' in status:
        try:
            tour_number = df[(df.Tour_Type == 'Day_Drive') &
                             ((df.Date - current_date) >= datetime.timedelta(days=-1))].index[attempt_number - 1]
        except IndexError:
            print('Couldn\'t find correct tour')
            tour_number = df[(df.Tour_Type != 'Audition')].index[attempt_number - 1]
    else:
        print('Couldn\'t find correct tour')
        tour_number = df[(df.Tour_Type != 'Audition')].index[attempt_number - 1]
    pyautogui.doubleClick(x + 469, y + 67 + 13 * tour_number)

    # Checks if "You need to change sites" message comes up
    # Messed up here and had to make it confusing
    screen_shot = None
    while screen_shot == sc.no_popup or screen_shot is None:
        with mss.mss() as sct:
            monitor = {'top': 507, 'left': 941, 'width': 23, 'height': 13}
            im = sct.grab(monitor)
            screen_shot = str(mss.tools.to_png(im.rgb, im.size))
    if screen_shot != sc.no_incorrect_site:
        pyautogui.click(m2['yes_change_sites'])


# m3 functions
def count_accommodations():
    sc.get_m3_coordinates()
    number_of_accommodations = 0
    number_of_canceled_accommodations = 0
    x, y = m3['title']
    while True:
        screenshot = cf.take_screenshot(x + 330, y + 64, 97, 7)
        if screenshot == sc.no_accommodations:
            return number_of_accommodations, number_of_canceled_accommodations
        else:
            screenshot_2 = cf.take_screenshot(x + 211, y + 66, 52, 5)
            if screenshot_2 == sc.canceled_accommodation:
                number_of_canceled_accommodations += 1
                y += 13
            else:
                number_of_accommodations += 1
                y += 13


def check_tour_type(number_of_tours, status):
    global errors
    sc.get_m3_coordinates()
    x, y = m3['title']
    tour_types_dict = cf.read_pickle_file('m3_tour_type.p')
    tour_type = tour_types_dict[cf.take_screenshot(x + 36, y + 143, 89, 12)]
    if 'u' in status and tour_type != 'Minivac':
        cf.print_colored_text('Can\'t upgrade day drive', 'red')
        errors += 1
    if 't' in status and tour_type != 'Day_Drive':
        cf.print_colored_text('TAVS are only for Day Drives.', 'red')
        errors += 1
    if (tour_type == 'Day_Drive' or tour_type == 'Canceled' or tour_type == 'Open_Reservation') and number_of_tours > 0:
        cf.print_colored_text(tour_type + ' - ' + str(number_of_tours), 'red')
        errors += 1
    elif tour_type == 'Minivac' and number_of_tours < 1:
        cf.print_colored_text(tour_type + ' - ' + str(number_of_tours), 'red')
        errors += 1
    else:
        log.info(tour_type + ' - ' + str(number_of_tours))
        cf.print_colored_text(tour_type + ' - ' + str(number_of_tours), 'green')
    return tour_type


def count_deposits():
    sc.get_m3_coordinates()
    number_of_deposits = 0
    pyautogui.click(m3['tour_packages'])
    x, y = m3['title']
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\balance.png',
                                           region=(700, 245, 850, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\balance.png',
                                               region=(700, 245, 850, 566))
    while True:
        # Counts number of deposits.
        # Breaks 'while' loop once a returned screenshot is blank
        deposit_screenshot = cf.take_screenshot(x + 255, y + 69, 6, 9)
        if deposit_screenshot == sc.no_deposits:
            return number_of_deposits
        else:
            number_of_deposits += 1
            y += 13


def create_deposit_dataframe():
    sc.get_m3_coordinates()
    x, y = m3['deposit_1']
    d = []
    number_of_deposits = count_deposits()
    if number_of_deposits == 0:
        number_of_refundable_deposits = 0
        return 'No Deposits', number_of_refundable_deposits
    for i in range(number_of_deposits):
        clipboard.copy('bad')
        pyautogui.click(x, y)
        y += 13
        pyautogui.click(m3['change_deposit'])
        sc.get_m6_coordinates()
        pyautogui.click(m6['description'])
        keyboard.send('ctrl + z')
        keyboard.send('ctrl + c')
        r = Tk()
        result = r.selection_get(selection="CLIPBOARD")
        while result == 'bad':
            pyautogui.click(m6['description'])
            keyboard.send('ctrl + z')
            keyboard.send('ctrl + c')
            result = r.selection_get(selection="CLIPBOARD")
        pyautogui.click(m6['view'])
        item_in_deposit = sc.get_m7_coordinates()
        if item_in_deposit is None:
            price = 0
        else:
            pyautogui.doubleClick(m7['amount'])
            time.sleep(0.5)
            keyboard.send('ctrl + c')
            r = Tk()
            price = str(r.selection_get(selection="CLIPBOARD").replace('-', ''))
            price = price.replace('.00', '')
        if 'refunded' in result.lower():
            deposit_type = 'Refunded'
            price = '0'
        elif 'minivac' in result.lower() or 'apply' in result.lower():
            deposit_type = 'Non_Refundable'
        elif 'ref' in result.lower():
            deposit_type = 'Refundable'
        else:
            deposit_type = 'Non_Refundable'
        d.append({'Deposit_Type': deposit_type, 'Price': price})
        pyautogui.click(m7['cancel'])
        pyautogui.click(m6['ok'])
    df = pd.DataFrame(d)  # Turn d into a dataframe
    deposit_df = df[['Deposit_Type', 'Price']]  # Reorders the columns in the dataframe.
    number_of_refundable_deposits = len(df[(df.Deposit_Type == 'Refundable')])
    return deposit_df, number_of_refundable_deposits


"""def count_premiums():
    date_1 = '1'
    date_2 = '1'
    date_3 = '2005'
    m = 0
    date_dictionary = {}
    while int(date_2) != 32:
        print('{} - {} - {}'.format(date_1, date_2, date_3))
        f = open('text_files\\dates.p', 'rb')
        date_dictionary = pickle.load(f)
        f.close()
        screenshot = cf.take_screenshot(958, 367, 13, 10)
        date_dictionary[str(screenshot)] = 'Error'
        f = open('text_files\\dates.p', 'wb')
        pickle.dump(date_dictionary, f)
        f.close()
        screenshot = cf.take_screenshot(284, 139 + 13 * m, 13, 10)
        screenshot_2 = cf.take_screenshot(299, 139 + 13 * m, 15, 10)
        screenshot_3 = cf.take_screenshot(315, 139 + 13 * m, 27, 10)
        date_dictionary[str(screenshot)] = date_1
        date_dictionary[str(screenshot_2)] = date_2
        date_dictionary[str(screenshot_3)] = date_3
        f = open('text_files\\dates.p', 'wb')
        pickle.dump(date_dictionary, f)
        f.close()
        m += 1
        if int(date_1) <= 8:
            date_1 = int(date_1)
            date_1 += 1
            date_1 = str(date_1)
        date_2 = int(date_2)
        date_2 += 1
        date_2 = str(date_2)
        if int(date_3) <= 2019:
            date_3 = int(date_3)
            date_3 += 1
            date_3 = str(date_3)"""


def add_premium_to_dictionary():
    global premium_dict
    sc.get_m3_coordinates()
    number_of_premiums = 0
    screenshot_number = 0
    pyautogui.click(m3['premiums'])
    x, y = m3['premium_1']
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\issued.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\issued.png',
                                               region=(514, 245, 889, 566))
    while True:
        screenshot = cf.take_screenshot(x - 223, y - 4, 90, 9)
        screenshot_number += 1
        try:
            screenshot = premium_dict[screenshot]
        except KeyError:
            pyautogui.click(x - 223, y - 4)
            pyautogui.click(m3['change_premium'])
            sc.get_m10_coordinates()
            time.sleep(1)
            pyautogui.doubleClick(m10['name'])
            keyboard.send('ctrl + z')
            keyboard.send('ctrl + c')
            # pyautogui.click(m3['change_premium'])
            copied_text = str(clipboard.paste())
            premium_dict[str(screenshot)] = copied_text
            f = open('text_files\\premiums.p', 'wb')
            pickle.dump(premium_dict, f)
            f.close()
            f = open('text_files\\premiums.p', 'rb')
            premium_dict = pickle.load(f)
            f.close()
            f = open('text_files\\premiums_backup.p', 'ab')
            pickle.dump(premium_dict, f)
            f.close()
            pyautogui.click(m10['ok'])
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changingpremium.png', region=(514, 245, 889, 566))
            while image is None:
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\issued.png',
                                                       region=(514, 245, 889, 566))
            time.sleep(0.5)
        if screenshot == 'Nothing':
            return number_of_premiums
        else:
            number_of_premiums += 1
            y += 13


def count_items_in_deposit():
    sc.get_m6_coordinates()
    number_of_deposit_items = 0
    x, y = m6['title']
    while True:
        screenshot = cf.take_screenshot(x + 339, y + 189, 10, 8)
        if screenshot == sc.no_deposit_items:
            return number_of_deposit_items
        else:
            number_of_deposit_items += 1
            y += 13


def apply_to_mv(deposit_df):
    pyautogui.click(m3['tour_packages'])
    pyautogui.click(m3['deposit_1'])
    pyautogui.click(m3['change_deposit'])
    deposit_item_amount = count_items_in_deposit()
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
    new_reference = old_reference.replace("D-", "U-")
    clipboard.copy(str(new_reference))
    pyautogui.click(m7['cancel'])
    sc.get_m6_coordinates()
    pyautogui.click(m6['description'])
    keyboard.send('ctrl + z')
    keyboard.send('ctrl + c')
    r = Tk()
    old_description = r.selection_get(selection="CLIPBOARD")
    if 'AMS' in old_description:
        keyboard.write('AMS/Minivac')
    pyautogui.click(m6['payment'])
    sc.get_m8_coordinates()
    pyautogui.click(m8['transaction_code'])
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\apply_to_mv.png',
                                           region=(136, 652, 392, 247))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\apply_to_mv.png',
                                               region=(136, 652, 392, 247))
    pyautogui.click(image)
    pyautogui.click(m8['reference'])
    keyboard.write('APPLY TO MV')
    pyautogui.click(m8['ok'])
    time.sleep(0.3)
    pyautogui.click(880, 565)
    pyautogui.click(m6['payment'])
    sc.get_m8_coordinates()
    time.sleep(0.3)
    pyautogui.click(m8['transaction_code'])
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_payment.png',
                                           region=(136, 652, 392, 247))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_payment.png',
                                               region=(136, 652, 392, 247))
    pyautogui.click(image)
    pyautogui.doubleClick(m8['amount'])
    keyboard.write(deposit_df.Price[0])
    pyautogui.click(m8['reference'])
    keyboard.write(new_reference)
    cf.pause('Ok?')
    pyautogui.click(m8['ok'])
    sc.get_m6_coordinates()
    pyautogui.click(m6['ok'])
    print('Applied Refundable Deposit to Minivac')
    deposit_df.Deposit_Type[0] = 'Non_Refundable'


def read_premiums(number_of_refundable_deposits):
    global premium_dict
    list_of_premiums = []
    number_of_dep_premiums = 0
    number_of_premiums = 0
    sc.get_m3_coordinates()
    pyautogui.click(m3['premiums'])
    x, y = m3['premium_1']
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\issued.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\issued.png',
                                               region=(514, 245, 889, 566))
    while True:
        screenshot = cf.take_screenshot(x - 223, y - 4, 90, 9)
        screenshot_2 = cf.take_screenshot(x - 129, y - 4, 8, 7)
        try:
            screenshot = premium_dict[screenshot]
        except KeyError:
            add_premium_to_dictionary()
            read_premiums(number_of_refundable_deposits)
        if screenshot == 'Nothing':
            break
        elif screenshot_2 == sc.canceled_premium:
            y += 13
        else:
            if '20' in screenshot or '40' in screenshot or '50 ' in screenshot or '99' in screenshot:
                if 'Live' in screenshot:
                    pass
                else:
                    number_of_dep_premiums += 1
            if '20' in screenshot:
                screenshot = '20'
            elif '40' in screenshot:
                screenshot = '40'
            elif '$50 ' in screenshot:
                screenshot = '50'
            elif '99' in screenshot:
                screenshot = '99'
            list_of_premiums.append(screenshot)
            number_of_premiums += 1
            y += 13
    if number_of_dep_premiums != number_of_refundable_deposits:
        log.info(str(number_of_dep_premiums) + ' DEP Premium(s) - ' +
                 str(number_of_refundable_deposits) + ' Refundable Deposit(s)')
        print(u"\u001b[31m" + str(number_of_dep_premiums) + ' DEP Premium(s) - ' +
              str(number_of_refundable_deposits) + ' Refundable Deposit(s)' + u"\u001b[0m")
    else:
        print(u"\u001b[32m" + str(number_of_dep_premiums) + ' DEP Premium(s) - ' +
              str(number_of_refundable_deposits) + ' Refundable Deposit(s)' + u"\u001b[0m")
    if number_of_premiums != len(set(list_of_premiums)):
        print(u"\u001b[31m" + str(number_of_premiums) + ' Premium(s) - DUPLICATES' + u"\u001b[0m")
    else:
        print(u"\u001b[32m" + str(number_of_premiums) + ' Premium(s) - No Duplicates' + u"\u001b[0m")
    return list_of_premiums


def check_for_dep_premium(deposit_df, premiums):
    sc.get_m3_coordinates()
    global errors
    for index, row in deposit_df.iterrows():
        if row['Deposit_Type'] == 'Refundable' and row['Price'] == '40':
            if '40' in premiums:
                print(u"\u001b[32m" + '$40 DEP is present' + u"\u001b[0m")
            else:
                print(u"\u001b[31m" + 'Missing $40 DEP' + u"\u001b[0m")
                errors += 1
        elif row['Deposit_Type'] == 'Refundable' and row['Price'] == '50':
            if '50' in premiums:
                print(u"\u001b[32m" + '$50 DEP is present' + u"\u001b[0m")
            else:
                print(u"\u001b[31m" + 'Missing $50 DEP' + u"\u001b[0m")
                errors += 1
        elif row['Deposit_Type'] == 'Refundable' and row['Price'] == '20':
            if '20' in premiums:
                print(u"\u001b[32m" + '$20 DEP is present' + u"\u001b[0m")
            else:
                print(u"\u001b[31m" + 'Missing $20 DEP' + u"\u001b[0m")
                errors += 1
        elif row['Deposit_Type'] == 'Refundable' and row['Price'] == '99':
            if '99' in premiums:
                print(u"\u001b[32m" + '$99 DEP is present' + u"\u001b[0m")
            else:
                print(u"\u001b[31m" + 'Missing $99 DEP' + u"\u001b[0m")
                errors += 1
        elif row['Deposit_Type'] == 'Refundable' and row['Price'] == '100':
            if '100' in premiums:
                print(u"\u001b[32m" + '$100 DEP is present' + u"\u001b[0m")
            else:
                print(u"\u001b[31m" + 'Missing $100 DEP' + u"\u001b[0m")
                errors += 1


def confirm_tour_status(status):
    """Checks that the tour status is correct"""
    global errors
    tour_status = None
    pyautogui.click(m3['tour'])
    pyautogui.click(m3['accommodations'])
    x, y = m3['title']
    tour_status = sc.m3_tour_status[cf.take_screenshot(x + 37, y + 170, 94, 11)]
    if status == 'c' and tour_status not in ['Confirmed', 'Showed', 'On_Tour', 'No_Show']:
        print(u"\u001b[33;1m" + 'TOUR STATUS MIGHT BE INCORRECT' + u"\u001b[0m")
        errors += 1
    elif status == 'r' and tour_status not in ['Rescheduled', 'No_Tour']:
        print(u"\u001b[33;1m" + 'TOUR STATUS MIGHT BE INCORRECT' + u"\u001b[0m")
        errors += 1
    elif status == 'x' and tour_status not in ['Canceled']:
        print(u"\u001b[33;1m" + 'TOUR STATUS MIGHT BE INCORRECT' + u"\u001b[0m")
        errors += 1
    elif status in ['c', 'r', 'x']:
        print(u"\u001b[32m" + 'Tour Status - ' + tour_status + u"\u001b[0m")
    return tour_status


def notes(status):
    global errors
    sc.get_m3_coordinates()
    pyautogui.click(m3['notes'])
    x, y = m3['notes']
    copied = []
    while True:
        pyautogui.click(x, y + 40)
        pyautogui.click(m3['notes_change'])
        attempts = 0
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changing_note.png',
                                               region=(514, 245, 889, 566))
        while attempts <= 3 and image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changing_note.png',
                                                   region=(514, 245, 889, 566))
            attempts += 1
        if image is not None:
            x_2, y_2 = image
            pyautogui.click(x_2 + 25, y_2 + 75)
            pyautogui.dragTo(x_2 + 250, y_2 + 150, button='left')
            keyboard.send('ctrl + c')  # Copy note
            r = Tk()
            result = r.selection_get(selection="CLIPBOARD")
            if result in copied:
                print(u"\u001b[31m" + 'COULDN\'T FIND CORRECT NOTE' + u"\u001b[0m")
                errors += 1
                pyautogui.click(x_2 + 200, y_2 + 250)
                return
            for key, value in sol_numbers.items():
                words = re.findall(r'\w+', key)
                if words[1].lower() in result.lower() and words[1].lower() != 'major':
                    print(u"\u001b[31m" + 'IMPORTANT NOTE' + u"\u001b[0m")
            if status == 'c' and 'conf' in result.lower():
                print(u"\u001b[32m" + 'Confirm note is present' + u"\u001b[0m")
                pyautogui.click(x_2 + 200, y_2 + 250)
                return
            elif status == 'x' and ('nq' in result.lower() or 'canc' in result.lower() or 'cxl' in result.lower()):
                print(u"\u001b[32m" + 'Cancel note is present' + u"\u001b[0m")
                pyautogui.click(x_2 + 200, y_2 + 250)
                if 'nq' in result.lower():
                    pyautogui.click(m3['tour'])
                    if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\tour_result.png',
                                                      region=(514, 400, 889, 500)) is None:
                        print(u"\u001b[32m" + 'Tour Result is correct' + u"\u001b[0m")
                    else:
                        print(u"\u001b[31m" + 'NO TOUR RESULT' + u"\u001b[0m")
                        errors += 1
                return
            elif status == 'r' and ('rxl' in result.lower() or 'open' in result.lower() or ' od ' in result.lower()):
                print(u"\u001b[32m" + 'Reschedule note is present' + u"\u001b[0m")
                pyautogui.click(x_2 + 200, y_2 + 250)
                return
            else:
                copied.append(result)
                pyautogui.click(x_2 + 200, y_2 + 250)
                y += 13
        else:
            print(u"\u001b[31m" + 'NO NOTES' + u"\u001b[0m")
            errors += 1
            return


def confirm_sol_in_userfields(sol, tour_status):
    if tour_status == 'Confirmed':
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
        while image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                   region=(514, 245, 889, 566))
        x, y = image
        pyautogui.click(x, y + 18)  # User Fields Tab
        if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_confirmer_sol.png',
                                          region=(514, 245, 889, 566)) is None:
            print(u"\u001b[32m" + 'Sol number is good' + u"\u001b[0m")
        else:
            pyautogui.doubleClick(x + 115, y + 222)
            keyboard.write(sol)
            print('Sol number was changed')
        pyautogui.click(x - 65, y + 18)


def enter_personnel(sol, status):
    sc.get_m3_coordinates()
    for i in status:
        pyautogui.click(m3['personnel'])
        pyautogui.click(m3['insert_personnel'])
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
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_personnel_titles_menu'
                                                   '.png', region=(514, 245, 889, 566))
        x_3, y_3 = image
        pyautogui.click(x_3 + 75, y_3 + 150)  # Close
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_addingrecord.png',
                                               region=(514, 245, 889, 566))
        while image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_addingrecord.png',
                                                   region=(514, 245, 889, 566))
        x_4, y_4 = image
        if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_confirmer.png',
                                          region=(514, 245, 889, 566)) is None:
            pyautogui.click(x_4 + 90, y_4 + 80)
            keyboard.write("cc")
        pyautogui.click(x_4 + 90, y_4 + 105)
        if i == 'c':
            keyboard.write("cc")
        elif i == 'r':
            keyboard.write("r")
        elif i == 'x':
            keyboard.write("c")
        elif i == 'u':
            keyboard.write("u")
        elif i == 't':
            keyboard.write("t")
        pyautogui.click(x_4 + 90, y_4 + 350)


def check_for_duplicate_personnel(df, status):
    sc.get_m3_coordinates()
    pyautogui.click(m3['title'])
    sc.get_m2_coordinates()
    select_tour(df, status)


def convert_excel_to_csv():
    xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\Downloads\\3.xlsx")
    df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
    df.to_csv('file.csv')


def mark_row_as_completed(index):
    wb = openpyxl.load_workbook(filename='C:\\Users\\Jared.Abrahams\\Downloads\\3.xlsx')
    ws = wb.worksheets[0]
    ws.cell(row=int(index) + 2, column=8).value = 'x'
    wb.save('C:\\Users\\Jared.Abrahams\\Downloads\\3.xlsx')


def count_pids():
    number_of_pids = 0
    with open('file.csv') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            number_of_pids += 1
    return number_of_pids


def show_progress(pid, progress, number_of_pids):
    percentage = round(progress * 100 / number_of_pids, 2)
    print('{} / {} - {} - {}'.format(str(progress), str(number_of_pids), str(percentage) + '%', str(pid)))


def assign_variables(row):
    global sol
    global progress
    progress = 0
    progress += 1
    status = []
    index = row['']
    pid = row['PID'].replace('.0', '')
    completed = row['Completed']
    if row['Sol'] != '':
        try:
            sol = sol_numbers[row['Sol']]
        except TypeError:
            print('Unrecognized Name - ' + sol)
            sol = 'SOL' + input('Type Sol number (just numbers):')
    elif row['Sol'] == '' and sol == 0:
        sys.exit('No Sol Number')
    if row['conf'] in ['x', 'X']:
        status.append('c')
    if row['rxl'] in ['x', 'X']:
        status.append('r')
    if row['cxl'] in ['x', 'X']:
        status.append('x')
    if row['ug'] in ['x', 'X']:
        status.append('u')
    if row['tav'] in ['x', 'X']:
        status.append('t')
    return index, pid, status, completed


def automatic_confirmation():
    global errors
    global sol
    errors = 0
    progress = 1
    convert_excel_to_csv()
    number_of_pids = count_pids()
    with open('file.csv') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            errors = 0
            index, pid, status, completed = assign_variables(row)
            if completed == 'x':
                progress += 1
                continue
            show_progress(pid, progress, number_of_pids)
            cf.search_pid(pid)
            cf.double_check_pid(pid)
            select_tour(status)
            number_of_tours, number_of_canceled_tours = count_accommodations()
            m3_tour_type, m3_tour_status, m3_tour_date = gather_m3_data()
            if (m3_tour_type == 'Open_Reservation' or m3_tour_type == 'No_Tour') and \
                    m3_tour_date != datetime.datetime.strptime('1/1/1900', "%m/%d/%Y"):
                print(u"\u001b[31m" + 'DATE IS INCORRECT' + u"\u001b[0m")
            tour_type = check_tour_type(number_of_tours, status)

            # Deposit Stuff
            deposit_df, number_of_refundable_deposits = create_deposit_dataframe()
            try:
                rows, columns = deposit_df.shape
                print(tabulate(deposit_df, headers='keys', tablefmt='psql'))
                if rows > 1 and deposit_df.Deposit_Type[0] == 'Refundable' and (deposit_df.Price[1] == '9' or
                                                                                deposit_df.Price[1] == '19' or
                                                                                deposit_df.Price[1] == '29'):
                    apply_to_mv(deposit_df)
                    number_of_refundable_deposits -= 1
                premiums = read_premiums(number_of_refundable_deposits)
                check_for_dep_premium(deposit_df, premiums)
            except AttributeError:
                rows, columns = 0, 0
                cf.print_colored_text('No deposits', 'green')
            try:
                ug = row['ug']
                if ug == "X" or ug == "x":
                    enter_personnel(sol, 'u')
            except KeyError:
                pass
            try:
                tav = row['tav']
                if tav == "X" or tav == "x":
                    enter_personnel(sol, 't')
            except KeyError:
                pass
            if 'c' in status and 'r' in status:
                tour_status = confirm_tour_status('c')
                notes('c')
                confirm_sol_in_userfields(sol, tour_status)
                enter_personnel(sol, 'c')
                enter_personnel(sol, 'r')
            elif 'r' in status and 'x' in status:
                tour_status = confirm_tour_status('x')
                notes('x')
                enter_personnel(sol, 'r')
                enter_personnel(sol, 'x')
            elif 'c' in status and 'r' not in status and 'x' not in status:
                tour_status = confirm_tour_status('c')
                notes('c')
                confirm_sol_in_userfields(sol, tour_status)
                enter_personnel(sol, 'c')
            elif 'r' in status and 'c' not in status and 'x' not in status:
                tour_status = confirm_tour_status('r')
                notes('r')
                enter_personnel(sol, 'r')
            elif 'x' in status and 'c' not in status and 'r' not in status:
                tour_status = confirm_tour_status('x')
                notes('x')
                enter_personnel(sol, 'x')
            #  check_for_duplicate_personnel(df, status)
            mark_row_as_completed(index)
            cf.pause("Everthing ok?")
            # if errors > 0 or tour_type == 'Minivac':
            # pause("Everything ok?")
            progress += 1
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
            errors = 0


"""number_dictionary = {}
x = 0
for i in range(11):
    screenshot = cf.take_screenshot(284, 140 + x * 13, 6, 9, True)
    if x < 10:
        number_dictionary[screenshot] = str(x)
    else:
        number_dictionary[screenshot] = 'nothing'
    x += 1
f = open('text_files\\numbers.p', 'wb')
pickle.dump(number_dictionary, f)
f.close()
with mss.mss() as sct:
    monitor = {'top': 555 - 0, 'left': 1489 - 6, 'width': 6, 'height': 9}
    im = sct.grab(monitor)
    screenshot = str(mss.tools.to_png(im.rgb, im.size))
    print(screenshot)
numbers = cf.read_pickle_file('numbers.p')
screenshot = cf.take_screenshot(1489 - 6, 555 - 0, 6, 9)
print(screenshot)
number = numbers[screenshot]
print(number)"""

if __name__ == "__main__":
    sol = 0
    log = logging.getLogger(__name__)
    log.setLevel(logging.DEBUG)
    formatter = logging.Formatter(fmt="%(asctime)s:%(filename)s:%(levelname)s:%(message)s",
                                  datefmt="%Y-%m-%d - %H:%M:%S")
    fh = logging.FileHandler("mylog.log")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)
    log.addHandler(fh)
    automatic_confirmation()
    """confirmation_sheet = ConfirmationSheet('3.xlsx')
    confirmation_sheet.convert_to_csv()
    number_of_pids = confirmation_sheet.count_pids()"""

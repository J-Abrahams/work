import importlib
import re
import csv
import time
from tkinter import Tk
import keyboard
import mss
import mss.tools
import pandas as pd
import pyautogui
import pyperclip
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8
import datetime

errors = 0


#  TODO Make it so the confirmer number only gets changed when the tour status is confirmed.
#  TODO If there is no item in a deposit, make it so the program doesn't hang on that deposit. 1398246
#  TODO If the description of a deposit is AMS/Refundable Deposit/Apply to Minivac, make it treat that as non-refundable
#  TODO If there is a canceled minivac, make the minivac - 0 message green instead of red.
#  TODO Check for accommodation cancel notes 1423766
#  TODO Partial refunds 1423766
#  TODO Check title of note when it's a cancel
#  TODO Fix duplicate deposits.
#  TODO If a deposit is refunded, don't create a missing dep message 1384377


def search_pid(pid_number):
    sc.get_m1_coordinates()
    pyautogui.doubleClick(m1['search'])
    keyboard.write(pid_number)
    pyautogui.click(m1['find_now'])
    pyautogui.click(m1['change'])


def double_check_pid(pid_number):
    sc.get_m2_coordinates()
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


def select_tour():
    sc.get_m2_coordinates()
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


def check_tour_for_error():
    sc.get_m3_coordinates()
    with mss.mss() as sct:
        x, y = m3['title']
        monitor = {'top': y + 171, 'left': x + 40, 'width': 52, 'height': 12}
        im = sct.grab(monitor)
        tour_status = str(mss.tools.to_png(im.rgb, im.size))
        if tour_status == sc.error:
            input('Is this the correct tour?')


def count_accommodations():
    sc.get_m3_coordinates()
    number_of_accommodations = 0
    number_of_canceled_accommodations = 0
    x, y = m3['title']
    while True:
        with mss.mss() as sct:
            monitor = {'top': y + 64, 'left': x + 330, 'width': 97, 'height': 7}
            im = sct.grab(monitor)
            screenshot = str(mss.tools.to_png(im.rgb, im.size))
        if screenshot == sc.no_accommodations:
            return number_of_accommodations, number_of_canceled_accommodations
        else:
            with mss.mss() as sct:
                monitor = {'top': y + 66, 'left': x + 211, 'width': 52, 'height': 5}
                im = sct.grab(monitor)
                screenshot_2 = str(mss.tools.to_png(im.rgb, im.size))
            if screenshot_2 == sc.canceled_accommodation:
                number_of_canceled_accommodations += 1
                y += 13
            else:
                number_of_accommodations += 1
                y += 13


def check_tour_type(number_of_tours):
    global errors
    sc.get_m3_coordinates()
    with mss.mss() as sct:
        x, y = m3['title']
        monitor = {'top': y + 143, 'left': x + 36, 'width': 89, 'height': 12}
        im = sct.grab(monitor)
        tour_type = sc.tour_type[str(mss.tools.to_png(im.rgb, im.size))]
    if (tour_type == 'day_drive' or tour_type == 'canceled' or tour_type == 'open_reservation') and number_of_tours > 0:
        print('\x1b[6;30;41m' + tour_type + ' - ' + str(number_of_tours) + '\x1b[0m')
        errors += 1
    elif tour_type == 'minivac' and number_of_tours < 1:
        print('\x1b[6;30;41m' + tour_type + ' - ' + str(number_of_tours) + '\x1b[0m')
        errors += 1
    else:
        print('\x1b[6;30;42m' + tour_type + ' - ' + str(number_of_tours) + '\x1b[0m')
    return tour_type


def check_for_refundable_deposit():
    sc.get_m3_coordinates()
    deposits = {}
    list_of_deposits = []
    list_of_refundable_deposits = []
    z = 1
    number_of_deposits = 0
    non_refundable_total = 0
    pyautogui.click(m3['tour_packages'])
    x, y = m3['title']
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\balance.png',
                                           region=(700, 245, 850, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\balance.png',
                                               region=(700, 245, 850, 566))
    while True:
        with mss.mss() as sct:
            monitor = {'top': y + 69, 'left': x + 255, 'width': 6, 'height': 9}
            im = sct.grab(monitor)
            screenshot = str(mss.tools.to_png(im.rgb, im.size))
        if screenshot == sc.no_deposits and number_of_deposits == 0:
            return non_refundable_total, list_of_refundable_deposits
        elif screenshot == sc.no_deposits:
            break
        else:
            number_of_deposits += 1
            y += 13
    x, y = m3['deposit_1']
    for i in range(number_of_deposits):
        pyautogui.click(x, y)
        y += 13
        pyautogui.click(m3['change_deposit'])
        sc.get_m6_coordinates()
        pyautogui.click(m6['description'])
        keyboard.send('ctrl + z')
        keyboard.send('ctrl + c')
        result = clipboard.paste()
        pyautogui.click(m6['view'])
        item_in_deposit = sc.get_m7_coordinates()
        if item_in_deposit is None:
            price = 0
        else:
            pyautogui.doubleClick(m7['amount'])
            time.sleep(0.5)
            keyboard.send('ctrl + c')
            price = str(clipboard.paste().replace('-', ''))
            price = price.replace('.00', '')
        if 'ref' in result.lower():
            deposits[0 + z] = ['refundable', price]
        else:
            deposits[0 + z] = ['non-refundable', price]
            # non_refundable_total += int(price)
        z += 1
        pyautogui.click(m7['cancel'])
        pyautogui.click(m6['ok'])
    print(deposits)
    return deposits
    """
    sc.get_m3_coordinates()
    deposits = {}
    list_of_deposits = []
    list_of_refundable_deposits = []
    screenshot = None
    z = 1
    number_of_deposits = 0
    non_refundable_total = 0
    keep_going = 1
    pyautogui.click(m3['tour_packages'])
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\balance.png',
                                           region=(700, 245, 850, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\balance.png',
                                               region=(700, 245, 850, 566))
    x, y = m3['title']
    while keep_going == 1:
        with mss.mss() as sct:
            monitor = {'top': y + 69, 'left': x + 268, 'width': 129, 'height': 7}
            im = sct.grab(monitor)
            screenshot = str(mss.tools.to_png(im.rgb, im.size))
        if screenshot == sc.no_deposits:
            keep_going = 0
        else:
            list_of_deposits.append(str(screenshot))
            number_of_deposits += 1
            y += 13
    x, y = m3['deposit_1']
    for i in range(number_of_deposits):
        pyautogui.click(x, y)
        y += 13
        pyautogui.click(m3['change_deposit'])
        sc.get_m6_coordinates()
        x_2, y_2 = m6['title']
        number_of_payments_in_deposit = -2
        with mss.mss() as sct:
            monitor = {'top': y_2 + 188, 'left': x_2 + 169, 'width': 49, 'height': 8}
            im = sct.grab(monitor)
            screenshot = str(mss.tools.to_png(im.rgb, im.size))
        while screenshot != sc.no_payments_in_deposit:
            with mss.mss() as sct:
                number_of_payments_in_deposit += 1
                y_2 += 13
                monitor = {'top': y_2 + 188, 'left': x_2 + 169, 'width': 49, 'height': 8}
                im = sct.grab(monitor)
                screenshot = str(mss.tools.to_png(im.rgb, im.size))
                print(screenshot)
        y_2 = y_2 + number_of_payments_in_deposit * 13
        # TODO Use screenshots to see if refundable or not
        pyautogui.click(m6['description'])
        keyboard.send('ctrl + z')
        keyboard.send('ctrl + c')
        r = Tk()
        result = r.selection_get(selection="CLIPBOARD")
        with mss.mss() as sct:
            monitor = {'top': y_2 + 188, 'left': x_2 + 186, 'width': 5, 'height': 10}
            im = sct.grab(monitor)
            screenshot = str(mss.tools.to_png(im.rgb, im.size))
            try:
                digit_1 = sc.deposit_amounts[screenshot]
            except KeyError:
                digit_1 = 0
                print(screenshot)
            monitor = {'top': y_2 + 188, 'left': x_2 + 192, 'width': 5, 'height': 10}
            im = sct.grab(monitor)
            screenshot = str(mss.tools.to_png(im.rgb, im.size))
            try:
                digit_2 = sc.deposit_amounts[screenshot]
            except KeyError:
                print(screenshot)
                digit_2 = 0
            monitor = {'top': y_2 + 188, 'left': x_2 + 198, 'width': 5, 'height': 10}
            im = sct.grab(monitor)
            screenshot = str(mss.tools.to_png(im.rgb, im.size))
            try:
                digit_3 = sc.deposit_amounts[screenshot]
            except KeyError:
                print(screenshot)
                digit_3 = 0
        deposit_price = int(str(digit_1) + str(digit_2) + str(digit_3))
        if 'ref' in result.lower():
            deposits[0 + z] = ['refundable', deposit_price]
        else:
            deposits[0 + z] = ['non-refundable', deposit_price]
            # non_refundable_total += int(price)
        z += 1
        sc.get_m7_coordinates()
        pyautogui.click(m7['cancel'])
        pyautogui.click(m6['ok'])
    print(deposits)
    return deposits"""


def apply_to_mv(deposits):
    try:
        deposit_1 = deposits[1]
        deposit_2 = deposits[2]
    except (KeyError, IndexError):
        return
    if 'refundable' in deposit_1 and ('9' in deposit_2 or '19' in deposit_2 or '29' in deposit_2):
        print('\x1b[6;30;41m' + 'Deposit needs to be changed' + '\x1b[0m')
        global errors
        errors += 1


def read_premiums():
    """Adds all premiums to the list premiums and checks if there are any duplicates among them"""
    global errors
    sc.get_m3_coordinates()
    list_of_premiums = []
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
            list_of_premiums.append(str(mss.tools.to_png(im.rgb, im.size)))
        y += 13
        pyautogui.click(x, y)
        time.sleep(0.3)
        is_premium_blue = pyautogui.pixelMatchesColor(x, y, (8, 36, 107))
        if is_premium_blue is False:
            is_premium_blue = pyautogui.pixelMatchesColor(x, y, (8, 36, 107))
    if len(list_of_premiums) != len(set(list_of_premiums)):
        print('\x1b[6;30;41m' + str(number_of_premiums) + ' Premiums - DUPLICATES' + '\x1b[0m')
        errors += 1
    else:
        print('\x1b[6;30;42m' + str(number_of_premiums) + ' Premiums - No Duplicates' + '\x1b[0m')
    return list_of_premiums


def check_for_dep_premium(deposits, premiums):
    sc.get_m3_coordinates()
    global errors
    try:
        for value in deposits.values():
            if value[0] == 'refundable' and value[1] == '40':
                if any(sc.dep_40_cc in s for s in premiums) or any(sc.dep_40_cash in s for s in premiums) or \
                        any(sc.d40_cc_dep in s for s in premiums) or any(sc.d40_dep in s for s in premiums):
                    print('\x1b[6;30;42m' + '$40 DEP is present' + '\x1b[0m')
                else:
                    print('\x1b[6;30;41m' + 'Missing $40 DEP' + '\x1b[0m')
                    errors += 1
            elif value[0] == 'refundable' and value[1] == '50':
                if any(sc.dep_50_cc in s for s in premiums) or any(sc.dep_50_cash in s for s in premiums) or \
                        any(sc.d50_cc_dep in s for s in premiums):
                    print('\x1b[6;30;42m' + '$50 DEP is present' + '\x1b[0m')
                else:
                    print('\x1b[6;30;41m' + 'Missing $50 DEP' + '\x1b[0m')
                    errors += 1
            elif value[0] == 'refundable' and value[1] == '20':
                if any(sc.d20_cc_dep in s for s in premiums) or any(sc.dep_20_cc in s for s in premiums):
                    print('\x1b[6;30;42m' + '$20 DEP is present' + '\x1b[0m')
                else:
                    print('\x1b[6;30;41m' + 'Missing $20 DEP' + '\x1b[0m')
                    errors += 1
            elif value[0] == 'refundable' and value[1] == '99':
                if any(sc.dep_99_cc in s for s in premiums):
                    print('\x1b[6;30;42m' + '$99 DEP is present' + '\x1b[0m')
                else:
                    print('\x1b[6;30;41m' + 'Missing $99 DEP' + '\x1b[0m')
                    errors += 1
    except AttributeError:
        print('\x1b[6;30;42m' + 'No deposits' + '\x1b[0m')


def confirm_tour_status(status):
    """Checks that the tour status is correct"""
    global errors
    tour_status = None
    pyautogui.click(m3['tour'])
    pyautogui.click(m3['accommodations'])
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
        else:
            tour_status = 'confirmed'
            print('\x1b[6;30;42m' + 'Tour status - Confirmed' + '\x1b[0m')
            return tour_status
        if image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\on_tour.png',
                                                   region=(x + 27, y + 132, 131, 103))
        else:
            tour_status = 'showed'
            print('\x1b[6;30;42m' + 'Tour status - Showed' + '\x1b[0m')
            return tour_status
        if image is None:
            print('\x1b[6;30;41m' + 'TOUR STATUS MIGHT BE INCORRECT' + '\x1b[0m')
            errors += 1
        else:
            tour_status = 'on_tour'
            print('\x1b[6;30;42m' + 'Tour status - On Tour' + '\x1b[0m')
            return tour_status
    elif status == 'r':
        if pyautogui.pixelMatchesColor(x + 48, y + 171, (0, 0, 0)) is True:
            tour_status = 'rescheduled'
            print('\x1b[6;30;42m' + 'Tour status - Rescheduled' + '\x1b[0m')
        elif pyautogui.pixelMatchesColor(x + 99, y + 171, (0, 0, 0)) is True:
            tour_status = 'Open'
            print('\x1b[6;30;42m' + 'Tour status - Open' + '\x1b[0m')
        else:
            print('\x1b[6;30;41m' + 'TOUR STATUS MIGHT BE INCORRECT' + '\x1b[0m')
            errors += 1
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
            errors += 1
        else:
            tour_status = 'canceled'
            print('\x1b[6;30;42m' + 'Tour status - Canceled' + '\x1b[0m')
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
            result = clipboard.paste()
            if result in copied:
                print('\x1b[6;30;41m' + 'COULDN\'T FIND CORRECT NOTE' + '\x1b[0m')
                errors += 1
                pyautogui.click(x_2 + 200, y_2 + 250)
                return
            if status == 'c' and 'conf' in result.lower():
                print('\x1b[6;30;42m' + 'Confirm note is present' + '\x1b[0m')
                pyautogui.click(x_2 + 200, y_2 + 250)
                return
            elif status == 'x' and ('nq' in result.lower() or 'canc' in result.lower() or 'cxl' in result.lower()):
                print('\x1b[6;30;42m' + 'Cancel note is present' + '\x1b[0m')
                pyautogui.click(x_2 + 200, y_2 + 250)
                if 'nq' in result.lower():
                    pyautogui.click(m3['tour'])
                    if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\tour_result.png',
                                                      region=(514, 400, 889, 500)) is None:
                        print('\x1b[6;30;42m' + 'Tour Result is correct' + '\x1b[0m')
                    else:
                        print('\x1b[6;30;41m' + 'NO TOUR RESULT' + '\x1b[0m')
                        errors += 1
                return
            elif status == 'r' and ('rxl' in result.lower() or 'open' in result.lower()):
                print('\x1b[6;30;42m' + 'Reschedule note is present' + '\x1b[0m')
                pyautogui.click(x_2 + 200, y_2 + 250)
                return
            else:
                copied.append(result)
                pyautogui.click(x_2 + 200, y_2 + 250)
                y += 13
        else:
            print('\x1b[6;30;41m' + 'NO NOTES' + '\x1b[0m')
            errors += 1
            return


def confirm_sol_in_userfields(sol, tour_status):
    if tour_status == 'confirmed':
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
        while image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                   region=(514, 245, 889, 566))
        x, y = image
        pyautogui.click(x, y + 18)  # User Fields Tab
        if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_confirmer_sol.png',
                                          region=(514, 245, 889, 566)) is None:
            print('\x1b[6;30;42m' + 'Sol number is good' + '\x1b[0m')
        else:
            pyautogui.doubleClick(x + 115, y + 222)
            pyperclip.copy(sol)
            keyboard.press_and_release('ctrl + v')
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


def convert_excel_to_csv():
    xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\Downloads\\3.xlsx")
    df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
    df.to_csv('file.csv')


def activation_sheet():
    for pid in pids:
        if pid != '':
            search_pid(pid)
            select_tour()
            read_premiums()
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


def automatic_confirmation():
    global errors
    convert_excel_to_csv()
    number_of_pids = 0
    progress = 1
    with open('file.csv') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            number_of_pids += 1
    with open('file.csv') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            pids = row['PID'].replace('.0', '')
            conf = row['conf']
            cxl = row['cxl']
            rxl = row['rxl']
            errors = 0
            print(str(progress) + '/' + str(number_of_pids))
            search_pid(pids)
            double_check_pid(pids)
            select_tour()
            check_tour_for_error()
            number_of_tours, number_of_canceled_tours = count_accommodations()
            tour_type = check_tour_type(number_of_tours)
            deposits = check_for_refundable_deposit()
            apply_to_mv(deposits)
            premiums = read_premiums()
            check_for_dep_premium(deposits, premiums)
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
            if (conf == "X" or conf == "x") and (rxl == "X" or rxl == "x"):
                tour_status = confirm_tour_status('c')
                notes('c')
                confirm_sol_in_userfields(sol, tour_status)
                enter_personnel(sol, 'c')
                enter_personnel(sol, 'r')
            elif (rxl == "X" or rxl == "x") and (cxl == "X" or cxl == "x"):
                tour_status = confirm_tour_status('x')
                notes('x')
                enter_personnel(sol, 'r')
                enter_personnel(sol, 'x')
            elif conf == "X" or conf == "x":
                tour_status = confirm_tour_status('c')
                notes('c')
                confirm_sol_in_userfields(sol, tour_status)
                enter_personnel(sol, 'c')
            elif rxl == "X" or rxl == "x":
                tour_status = confirm_tour_status('r')
                notes('r')
                enter_personnel(sol, 'r')
            elif cxl == "X" or cxl == "x":
                tour_status = confirm_tour_status('x')
                notes('x')
                enter_personnel(sol, 'x')
            if errors > 0 or tour_type == 'Minivac':
                input("Everything ok?")
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


print("Justin Locke - 4967 \nBrian Bennett - 3055")
sol = "SOL" + input("SOL #:")
automatic_confirmation()

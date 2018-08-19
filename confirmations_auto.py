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
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8
import datetime
from tabulate import tabulate
import sys

# import importlib
# importlib.reload(sc)
errors = 0


#  TODO Make it so the confirmer number only gets changed when the tour status is confirmed.
#  TODO If there is no item in a deposit, make it so the program doesn't hang on that deposit. 1398246
#  TODO If the description of a deposit is AMS/Refundable Deposit/Apply to Minivac, make it treat that as non-refundable
#  TODO Check for accommodation cancel notes 1423766
#  TODO Partial refunds 1423766
#  TODO Check title of note when it's a cancel
#  TODO Fix duplicate deposits.
#  TODO If a deposit is refunded, don't create a missing dep message 1384377
#  TODO Make program recognize CS $50 CC Deposit 1005465
#  TODO If an accommodation is canceled, check if the cancel box is ticked 1218776
#  TODO Make program recognize an Additional Nights deposit 984256
#  TODO If there are 2 refundable deposits and 2 dep premiums, don't count the 2 dep premiums as duplicates 1408810
#  TODO Fix data frame on PID 1425576
#  TODO If there are a lot of slashes in a note, count that as a confirm note.
#  TODO Automatically fill in the tour result when it's a cancel.

def take_screenshot(y, x, width, height, save_file=False):
    with mss.mss() as sct:
        monitor = {'top': y, 'left': x, 'width': width, 'height': height}
        im = sct.grab(monitor)
        screenshot = str(mss.tools.to_png(im.rgb, im.size))
        if save_file:
            now = datetime.datetime.now()
            output = now.strftime("%m-%d-%H-%M-%S.png".format(**monitor))
            mss.tools.to_png(im.rgb, im.size, output=output)
        return screenshot


def pause(message):
    print(message)
    while pyautogui.position() != (0, 1079):
        pass


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
    for i in range(3):
        if copied_text != pid_number:
            time.sleep(0.3)
            pyautogui.doubleClick(m2['prospect_id'])
            keyboard.send('ctrl + c')
            copied_text = clipboard.paste()
    if copied_text != pid_number:
        pause('Is the pid correct?')
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
        pause('Is the pid correct?')
        return


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
        tour_date = take_screenshot(y + 63, x + 330, 52, 10)
        tour_type = take_screenshot(y + 63, x + 402, 14, 10)
        tour_status = take_screenshot(y + 63, x + 484, 14, 10)
        try:
            tour_date = sc.dates[tour_date]
            if tour_date != 'Nothing':
                tour_date = datetime.datetime.strptime(tour_date, "%m/%d/%y")
        except KeyError:
            tour_date = None
        try:
            tour_type = sc.m2_tour_types[tour_type]
        except KeyError:
            print('Unrecognized tour type')
            print(tour_type)
            tour_type = None
        try:
            tour_status = sc.m2_tour_status[tour_status]
        except KeyError:
            print('Unrecognized tour status')
            print(tour_status)
            tour_status = None
        y += 13
        if tour_date != 'Nothing':
            try:
                # Where the screenshots get turned into dictionaries.
                d.append({'Date': tour_date, 'Tour_Type': tour_type, 'Tour_Status': tour_status})
            except NameError:
                pass
    df = pd.DataFrame(d)  # Turn d into a dataframe
    # df['Date'] = df['Date'].apply(lambda x: x.strftime('%Y-%m-%d') if not pd.isnull(x) else '')
    df = df[['Date', 'Tour_Type', 'Tour_Status']]  # Reorders the columns in the dataframe.
    print(tabulate(df, headers='keys', tablefmt='psql'))
    return df


def select_tour(df, status, attempt_number=1):
    x, y = m2['title']
    now = datetime.datetime.now()
    now = now.strftime("%m/%d/%y")
    current_date = datetime.datetime.strptime(now, "%m/%d/%y")
    # Returns the top tour that is Showed, not an Audition, and at most a week before the date we entered.
    # tour_number is the index of the correct tour. Ex: 1 if the second tour is the correct one.
    if 'c' in status:
        try:
            tour_number = df[((df.Tour_Status == 'Showed') | (df.Tour_Status == 'Confirmed') |
                              (df.Tour_Status == 'No_Show') | (df.Tour_Status == 'On_Tour')) &
                             (df.Tour_Type != 'Audition') &
                             ((df.Date - current_date) >= datetime.timedelta(days=-1))].index[attempt_number - 1]
        except IndexError:
            print('Couldn\'t find correct tour')
            tour_number = df[(df.Tour_Type != 'Audition')].index[attempt_number - 1]
    elif 'r' in status:
        try:
            tour_number = df[((df.Tour_Status == 'Rescheduled') &
                              ((df.Date - current_date) >= datetime.timedelta(days=0))) |
                             ((df.Tour_Type == 'Open_Reservation') &
                              (df.Date == datetime.datetime.strptime('1/01/00', "%m/%d/%y"))) &
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
            pause('Is this the correct tour?')


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
    if (tour_type == 'Day_Drive' or tour_type == 'Canceled' or tour_type == 'Open_Reservation') and number_of_tours > 0:
        print(u"\u001b[31m" + tour_type + ' - ' + str(number_of_tours) + u"\u001b[0m")
        errors += 1
    elif tour_type == 'Minivac' and number_of_tours < 1:
        print(u"\u001b[31m" + tour_type + ' - ' + str(number_of_tours) + u"\u001b[0m")
        errors += 1
    else:
        print(u"\u001b[32m" + tour_type + ' - ' + str(number_of_tours) + u"\u001b[0m")
    return tour_type


def read_deposits():
    sc.get_m3_coordinates()
    d = []
    number_of_deposits = 0
    number_of_refundable_deposits = 0
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
        deposit_screenshot = take_screenshot(y + 69, x + 255, 6, 9)
        if deposit_screenshot == sc.no_deposits and number_of_deposits == 0:
            return 'No Deposits', number_of_refundable_deposits
        elif deposit_screenshot == sc.no_deposits:
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
        r = Tk()
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


def apply_to_mv(deposit_df):
    pyautogui.click(m3['tour_packages'])
    pyautogui.click(m3['deposit_1'])
    pyautogui.click(m3['change_deposit'])
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
    pause('Ok?')
    pyautogui.click(m8['ok'])
    sc.get_m6_coordinates()
    pyautogui.click(m6['ok'])
    print('Applied Refundable Deposit to Minivac')
    deposit_df.Deposit_Type[0] = 'Non_Refundable'


def read_premiums(number_of_refundable_deposits):
    """Adds all premiums to the list premiums and checks if there are any duplicates among them"""
    global errors
    sc.get_m3_coordinates()
    list_of_premiums = []
    number_of_dep_premiums = 0
    pyautogui.click(m3['premiums'])
    time.sleep(0.3)
    pyautogui.click(m3['premium_1'])
    x, y = m3['premium_1']
    number_of_premiums = 0
    for i in range(8):
        with mss.mss() as sct:
            monitor = {'top': y - 4, 'left': x - 223, 'width': 90, 'height': 9}
            y += 13
            now = datetime.datetime.now()
            output = now.strftime("%m-%d-%H-%M-%S-%f.png".format(**monitor))
            sct_img = sct.grab(monitor)
            image = str(mss.tools.to_png(sct_img.rgb, sct_img.size))
            if image not in open('premiums.txt').read():
                mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)
                with open('premiums.txt', 'a') as out:
                    out.write('{},{}\n'.format(output, image))
    x, y = m3['premium_1']
    is_premium_blue = pyautogui.pixelMatchesColor(x, y, (8, 36, 107))
    while is_premium_blue is True:
        number_of_premiums += 1
        with mss.mss() as sct:
            monitor = {'top': y - 4, 'left': x - 223, 'width': 90, 'height': 9}
            im = sct.grab(monitor)
            screenshot = str(mss.tools.to_png(im.rgb, im.size))
            monitor = {'top': y - 4, 'left': x - 129, 'width': 8, 'height': 7}
            im = sct.grab(monitor)
            screenshot_2 = str(mss.tools.to_png(im.rgb, im.size))
            if screenshot_2 == sc.canceled_premium:
                screenshot += 'canceled'
                list_of_premiums.append(str(screenshot))
            else:
                try:
                    list_of_premiums.append(str(sc.dep_premiums[screenshot]))
                    number_of_dep_premiums += 1
                except KeyError:
                    list_of_premiums.append(str(screenshot))
        y += 13
        pyautogui.click(x, y)
        time.sleep(0.3)
        is_premium_blue = pyautogui.pixelMatchesColor(x, y, (8, 36, 107))
        if is_premium_blue is False:
            is_premium_blue = pyautogui.pixelMatchesColor(x, y, (8, 36, 107))
    if number_of_dep_premiums != number_of_refundable_deposits:
        print(u"\u001b[31m" + str(number_of_dep_premiums) + ' DEP Premium(s) - ' +
              str(number_of_refundable_deposits) + ' Refundable Deposits' + u"\u001b[0m")
    else:
        print(u"\u001b[32m" + str(number_of_dep_premiums) + ' DEP Premium(s) - ' +
              str(number_of_refundable_deposits) + ' Refundable Deposits' + u"\u001b[0m")
    if len(list_of_premiums) != len(set(list_of_premiums)):
        print(u"\u001b[31m" + str(number_of_premiums) + ' Premiums - DUPLICATES' + u"\u001b[0m")
        errors += 1
    else:
        print(u"\u001b[32m" + str(number_of_premiums) + ' Premiums - No Duplicates' + u"\u001b[0m")
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
            print(u"\u001b[32m" + 'Tour status - Confirmed' + u"\u001b[0m")
            return tour_status
        if image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\on_tour.png',
                                                   region=(x + 27, y + 132, 131, 103))
        else:
            tour_status = 'showed'
            print(u"\u001b[32m" + 'Tour status - Showed' + u"\u001b[0m")
            return tour_status
        if image is None:
            print(u"\u001b[33;1m" + 'TOUR STATUS MIGHT BE INCORRECT' + u"\u001b[0m")
            errors += 1
        else:
            tour_status = 'on_tour'
            print(u"\u001b[32m" + 'Tour status - On Tour' + u"\u001b[0m")
            return tour_status
    elif status == 'r':
        if pyautogui.pixelMatchesColor(x + 48, y + 171, (0, 0, 0)) is True:
            tour_status = 'rescheduled'
            print(u"\u001b[32m" + 'Tour status - Rescheduled' + u"\u001b[0m")
        elif pyautogui.pixelMatchesColor(x + 99, y + 171, (0, 0, 0)) is True:
            tour_status = 'Open'
            print(u"\u001b[32m" + 'Tour status - Open' + u"\u001b[0m")
        else:
            print(u"\u001b[33;1m" + 'TOUR STATUS MIGHT BE INCORRECT' + u"\u001b[0m")
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
            print(u"\u001b[33;1m" + 'TOUR STATUS MIGHT BE INCORRECT' + u"\u001b[0m")
            errors += 1
        else:
            tour_status = 'canceled'
            print(u"\u001b[32m" + 'Tour status - Canceled' + u"\u001b[0m")
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
    global sol
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
            status = []
            if row['Sol'] != '':
                sol = row['Sol'].replace('.0', '')
                sol = 'SOL' + sol
            elif row['Sol'] == '' and sol == 0:
                sys.exit('No Sol Number')
            pids = row['PID'].replace('.0', '')
            conf = row['conf']
            cxl = row['cxl']
            rxl = row['rxl']
            if conf == "X" or conf == "x":
                status.append('c')
            if rxl == "X" or rxl == "x":
                status.append('r')
            if cxl == "X" or cxl == "x":
                status.append('x')
            try:
                ug = row['ug']
                if ug == "X" or ug == "x":
                    status.append('u')
            except KeyError:
                pass
            try:
                tav = row['tav']
                if tav == "X" or tav == "x":
                    status.append('t')
            except KeyError:
                pass
            errors = 0
            print('{} / {} - {}'.format(str(progress), str(number_of_pids), str(pids)))
            search_pid(pids)
            double_check_pid(pids)
            df = create_data_frame()
            select_tour(df, status)
            check_tour_for_error()
            number_of_tours, number_of_canceled_tours = count_accommodations()
            tour_type = check_tour_type(number_of_tours)
            deposit_df, number_of_refundable_deposits = read_deposits()
            try:
                rows, columns = deposit_df.shape
            except AttributeError:
                rows, columns = 0, 0
            if rows > 1 and deposit_df.Deposit_Type[0] == 'Refundable' and (deposit_df.Price[1] == '9' or
                                                                            deposit_df.Price[1] == '19' or
                                                                            deposit_df.Price[1] == '29'):
                apply_to_mv(deposit_df)
                number_of_refundable_deposits -= 1
            if rows > 0:
                print(tabulate(deposit_df, headers='keys', tablefmt='psql'))
            premiums = read_premiums(number_of_refundable_deposits)
            if rows > 0:
                check_for_dep_premium(deposit_df, premiums)
            else:
                print(u"\u001b[32m" + 'No deposits' + u"\u001b[0m")
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
                pause("Everything ok?")
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
sol = 0
automatic_confirmation()

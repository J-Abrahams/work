import keyboard
import pyautogui
import time
import pyperclip
import sys
import datetime
from tkinter import Tk
import clipboard
import mss
import mss.tools
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8
import pandas as pd
import csv


def search_pid(pid_number):
    sc.get_m1_coordinates()
    pyautogui.doubleClick(m1['search'])
    keyboard.write(pid_number)
    pyautogui.click(m1['find_now'])
    pyautogui.click(m1['change'])


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


def double_check_pid(pid_number):
    sc.get_m2_coordinates()
    pyautogui.doubleClick(m2['prospect_id'])
    keyboard.send('ctrl + c')
    r = Tk()
    clipboard = r.selection_get(selection="CLIPBOARD")
    if clipboard != pid_number:
        input('Is the pid correct?')
        return
    if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\company.png',
                                      region=(514, 245, 889, 566)) is not None:
        return
    pyautogui.click(m2['company'])
    keyboard.send('ctrl + z')
    time.sleep(1)
    keyboard.send('ctrl + c')
    r = Tk()
    clipboard = r.selection_get(selection="CLIPBOARD")
    if 'pid' in clipboard.lower():
        input('Is the pid correct?')
        return


def change_deposit_title(price):
    """
    Checks to make sure that the deposit price is correct. Ex: If the sheet says $50, then this makes sure that the
    deposit is also $50.
    Changes the title from 'Refundable' to 'Refunded'
    """
    sc.get_m3_coordinates()
    amount = 0
    x, y = m3['deposit_1']
    while amount != price:
        pyautogui.click(m3['tour_packages'])
        pyautogui.click(x, y)
        pyautogui.click(m3['change_deposit'])
        sc.get_m6_coordinates()
        with mss.mss() as sct:
            x, y = m6['title']
            monitor = {'top': y + 187, 'left': x + 168, 'width': 51, 'height': 11}
            im = sct.grab(monitor)
            amount = sc.screenshot_dict[str(mss.tools.to_png(im.rgb, im.size))]
        if amount != price:
            pyautogui.click(m6['ok'])
            y += 13
    pyautogui.click(m6['description'])
    keyboard.send('ctrl + z')
    keyboard.send('ctrl + c')
    r = Tk()
    old_title = r.selection_get(selection="CLIPBOARD")
    new_title = old_title.replace("able", "ed")
    new_title = new_title.replace("ABLE", "ED")
    new_title = new_title.replace(" /", "/")
    new_title = new_title.replace("/ ", "/")
    keyboard.write(new_title)
    if 'prev' in new_title.lower():
        return 'prev'
    if new_title.lower() == 'ir/refunded deposit':
        return "ir"
    elif new_title.lower() == 'ams/refunded deposit':
        return "ams"
    elif new_title.lower() == 'sol/refunded deposit':
        return "sol"


def copy_reference_number():
    sc.get_m6_coordinates()
    pyautogui.click(m6['view'])
    sc.get_m7_coordinates()
    pyautogui.doubleClick(m7['reference'])
    keyboard.press_and_release('ctrl + c')
    time.sleep(0.5)
    old_reference = clipboard.paste()
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


"""def ams_cc_refund(date):
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
        return ams_credit_refund"""


def select_ams_refund_payment(date, price, description, reference_number=None):
    sc.get_m6_coordinates()
    attempts = 0
    global transaction_code
    if description == 'ams':
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
                transaction_code = 0
                sys.exit("Couldn't find correct choice")
    elif description == 'sol':
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
                sys.exit("Couldn't find correct choice")
    pyautogui.click(image)
    if (2 < transaction_code < 6) or transaction_code >= 7:
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


def select_ir_refund_payment(date, price, reference_number=None):
    sc.get_m6_coordinates()
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

"""def select_ams_refund_insert(x, y, price, date):
    pyautogui.doubleClick(x + 50, y + 190)  # Already entered payment
    time.sleep(1)
    pyautogui.doubleClick(x - 315, y + 160)  # Reference number
    keyboard.press_and_release('ctrl + c')
    time.sleep(1)
    old_reference = str(pyperclip.paste())

    if old_reference[0] != 'D':
        sys.exit("Wrong Reference")

    new_reference = old_reference.replace("D-", "R-")
    pyperclip.copy(new_reference)
    pyautogui.click(x - 150, y + 225)  # Exit out of the already entered payment
    pyautogui.click(x, y + 365)  # Insert button

    image = pyautogui.locateCenterOnScreen(
        'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\adding_tour_package_item.png', region=(136, 652, 392, 247))

    while image is None:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\adding_tour_package_item.png',
            region=(136, 652, 392, 247))

    x_1, y_1 = image
    pyautogui.click(x_1 + 100, y_1 + 65)  # Transaction code dropdown menu

    image = pyautogui.locateCenterOnScreen(
        'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(147, 566, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(147, 566, 889, 566))
    x_2, y_2 = image

    pyautogui.click(x_2, y_2)  # AMS refund
    pyautogui.doubleClick(x_1 + 85, y + 130)  # Amount
    keyboard.write(price)
    pyautogui.doubleClick(x_1 + 85, y_1 + 140)  # Reference
    keyboard.send('ctrl + v')
    pyautogui.doubleClick(x_1 + 85, y_1 + 165)  # Date
    keyboard.write(date)
    time.sleep(0.3)
    pyautogui.click(x_1 + 85, y_1 + 210)
    pyautogui.click(x_1 + 650, y_1 - 20)
    now = datetime.datetime.now()
    now_str = now.strftime("%Y-%m-%d-%H-%M-%S")
    outfile = pyautogui.screenshot(
        'C:\\Users\\Jared.Abrahams\\work\\deposit_screenshots\\ImageFile{}.png'.format(now_str))
    pyautogui.click(x + 150, y + 400)
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
    pyautogui.click(x - 20, y + 425)"""


"""def select_ams_cc_ref(x, y, price, date):
    pyautogui.doubleClick(x + 50, y + 190)  # Already entered payment
    time.sleep(1)
    pyautogui.doubleClick(x - 315, y + 160)  # Reference number
    keyboard.press_and_release('ctrl + c')
    time.sleep(1)
    old_reference = str(pyperclip.paste())

    if old_reference[0] != 'D':
        sys.exit("Wrong Reference")

    new_reference = old_reference.replace("D-", "R-")
    pyperclip.copy(new_reference)
    pyautogui.click(x - 150, y + 225)  # Exit out of the already entered payment
    pyautogui.click(x, y + 365)  # Insert button

    image = pyautogui.locateCenterOnScreen(
        'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\adding_tour_package_item.png',
        region=(147, 566, 889, 566))

    while image is None:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\adding_tour_package_item.png',
            region=(147, 566, 889, 566))

    x_1, y_1 = image
    pyautogui.click(x_1 + 100, y_1 + 65)  # Transaction code dropdown menu

    image = pyautogui.locateCenterOnScreen(
        'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_cc_ref.png', region=(147, 566, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_cc_ref.png', region=(147, 566, 889, 566))
    x_2, y_2 = image
    pyautogui.click(x_2, y_2)  # AMS refund
    pyautogui.doubleClick(x_1 + 85, y + 130)  # Amount
    keyboard.write(price)
    pyautogui.doubleClick(x_1 + 85, y_1 + 140)  # Reference
    keyboard.send('ctrl + v')
    pyautogui.doubleClick(x_1 + 85, y_1 + 165)  # Date
    keyboard.write(date)
    time.sleep(0.3)
    pyautogui.click(x_1 + 85, y_1 + 210)
    pyautogui.click(x_1 + 650, y_1 - 20)
    now = datetime.datetime.now()
    now_str = now.strftime("%Y-%m-%d-%H-%M-%S")
    outfile = pyautogui.screenshot(
        'C:\\Users\\Jared.Abrahams\\work\\deposit_screenshots\\ImageFile{}.png'.format(now_str))
    pyautogui.click(x + 150, y + 400)
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
    pyautogui.click(x - 20, y + 425)"""


"""def enter_rest_of_info(x, y, x_1, y_1):
    pyautogui.doubleClick(x_1 + 85, y + 130)  # Amount
    keyboard.write("50")
    pyautogui.doubleClick(x_1 + 85, y_1 + 140)  # Reference
    keyboard.send('ctrl + v')
    pyautogui.doubleClick(x_1 + 85, y_1 + 165)  # Date
    keyboard.write('7/10/18')
    time.sleep(0.3)
    pyautogui.click(x_1 + 85, y_1 + 210)
    pyautogui.click(x_1 + 650, y_1 - 20)
    now = datetime.datetime.now()
    now_str = now.strftime("%Y-%m-%d-%H-%M-%S")
    outfile = pyautogui.screenshot(
        'C:\\Users\\Jared.Abrahams\\work\\deposit_screenshots\\ImageFile{}.png'.format(now_str))
    pyautogui.click(x + 150, y + 400)
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
    pyautogui.click(x - 20, y + 425)"""


"""def ir_refund(x, y, price, date):
    pyautogui.doubleClick(x + 50, y + 190)  # Already entered payment
    time.sleep(1)
    pyautogui.doubleClick(x - 315, y + 160)  # Reference number
    keyboard.press_and_release('ctrl + c')
    time.sleep(1)
    old_reference = str(pyperclip.paste())

    if old_reference[0] != 'D':
        sys.exit("Wrong Reference")

    new_reference = old_reference.replace("D-", "R-")
    pyperclip.copy(new_reference)
    pyautogui.click(x - 150, y + 225)  # Exit out of the already entered payment
    pyautogui.click(x, y + 365)  # Insert button

    image = pyautogui.locateCenterOnScreen(
        'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\adding_tour_package_item.png',
        region=(147, 566, 889, 566))

    while image is None:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\adding_tour_package_item.png',
            region=(147, 566, 889, 566))

    x_1, y_1 = image
    pyautogui.click(x_1 + 100, y_1 + 65)  # Transaction code dropdown menu

    image = pyautogui.locateCenterOnScreen(
        'C:\\Users\\Jared.Abrahams\\Screenshots\\ir_refund.png', region=(147, 566, 889, 566))

    while image is None:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\ir_refund.png', region=(147, 566, 889, 566))

    x_2, y_2 = image
    pyautogui.click(x_2, y_2)  # IR refund
    pyautogui.doubleClick(x_1 + 85, y + 130)  # Amount
    keyboard.write(price)
    pyautogui.doubleClick(x_1 + 85, y_1 + 140)  # Reference
    keyboard.send('ctrl + v')
    pyautogui.doubleClick(x_1 + 85, y_1 + 165)  # Date
    keyboard.write(date)
    pyautogui.click(x_1 + 85, y_1 + 210)
    now = datetime.datetime.now()
    now_str = now.strftime("%Y-%m-%d-%H-%M-%S")
    time.sleep(0.3)
    outfile = pyautogui.screenshot(
        'C:\\Users\\Jared.Abrahams\\work\\deposit_screenshots\\ImageFile{}.png'.format(now_str))
    pyautogui.click(x + 150, y + 400)
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
    pyautogui.click(x - 20, y + 425)"""


"""def error(price):
    sc.get_m3_coordinates()
    amount = 0
    x, y = m3['deposit_1']
    while amount != price:
        pyautogui.click(m3['tour_packages'])
        pyautogui.click(x, y)
        pyautogui.click(m3['change_deposit'])
        sc.get_m6_coordinates()
        with mss.mss() as sct:
            x, y = m6['title']
            monitor = {'top': y + 187, 'left': x + 168, 'width': 51, 'height': 11}
            im = sct.grab(monitor)
            amount = sc.screenshot_dict[str(mss.tools.to_png(im.rgb, im.size))]
        if amount != price:
            pyautogui.click(m6['ok'])
            y += 13
    sc.get_m6_coordinates()
    pyautogui.click(m6['insert'])
    sc.get_m8_coordinates()
    pyautogui.click(m8['transaction_code'])
    pyautogui.click(m8['transaction_code_scroll_bar'])
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\void.png',
                                           region=(136, 652, 392, 247))
    while image is None:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\void.png',
            region=(136, 652, 392, 247))
    pyautogui.click(image)  # Error
    pyautogui.doubleClick(m8['amount'])
    keyboard.write("-50")
    #keyboard.write(price)
    pyautogui.doubleClick(m8['reference'])
    keyboard.write('ERROR')
    time.sleep(0.3)
    pyautogui.click(m8['ok'])
    time.sleep(0.3)
    pyautogui.click(880, 565)"""

def convert_excel_to_csv():
    xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\Downloads\\deposit_pids.xlsx")
    df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
    df.to_csv('file.csv')


def use_excel_sheet():
    convert_excel_to_csv()
    with open('file.csv') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            pids = row['PID'].replace('.0', '')
            search_pid(pids)
            select_tour()
            check_tour_for_error()
            deposit_type = change_deposit_title(price)
            if deposit_type != 'prev':
                copy_reference_number()
            else:
                ams_ir_or_sol = input('ams, ir, or sol:')
                reference_number = input('Reference Number starting with R:')
                if ams_ir_or_sol == 'ams':
                    select_ams_refund_payment(date, price, 'ams', reference_number)
                elif ams_ir_or_sol == 'ir':
                    select_ams_refund_payment(date, price, 'ams', reference_number)
                elif ams_ir_or_sol == 'sol':
                    select_ams_refund_payment(date, price, 'ams', reference_number)
            if deposit_type == 'ams':
                select_ams_refund_payment(date, price, 'ams')
            elif deposit_type == 'ir':
                select_ams_refund_payment(date, price, 'ir')
            elif deposit_type == 'sol':
                select_ams_refund_payment(date, price, 'sol')


pids = ['283614', '659413', '1064845', '1286522', '1303227', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '']
transaction_code = 0
date = input("Date")
price = input("Price")
print("Get PIDs from excel sheet?")
excel_sheet = input("(y) or (n)")
if excel_sheet != 'y':
    for pid in pids:
        if pid != '':
            search_pid(pid)
            select_tour()
            check_tour_for_error()
            error(price)
            deposit_type = change_deposit_title(price)
            if deposit_type != 'prev':
                copy_reference_number()
            else:
                ams_ir_or_sol = input('ams, ir, or sol:')
                reference_number = input('Reference Number starting with R:')
                if ams_ir_or_sol == 'ams':
                    select_ams_refund_payment(date, price, 'ams', reference_number)
                elif ams_ir_or_sol == 'ir':
                    select_ams_refund_payment(date, price, 'ams', reference_number)
                elif ams_ir_or_sol == 'sol':
                    select_ams_refund_payment(date, price, 'ams', reference_number)
            if deposit_type == 'ams':
                select_ams_refund_payment(date, price, 'ams')
            elif deposit_type == 'ir':
                select_ams_refund_payment(date, price, 'ir')
            elif deposit_type == 'sol':
                select_ams_refund_payment(date, price, 'sol')
else:
    use_excel_sheet()
'''search_pid('727085')
select_tour()
change_deposit_title()'''
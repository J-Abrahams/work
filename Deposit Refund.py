import keyboard
import pyautogui
import time
import pyperclip
import sys
import datetime


def search_for_tour_title():
    tour_menu = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
    while tour_menu is None:
        tour_menu = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                   region=(514, 245, 889, 566))
    x_1, y_1 = tour_menu
    return x_1, y_1


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
        audition = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\audition_2.png',
                                                  region=(514, 245, 889, 566))
        if audition is None:
            pyautogui.doubleClick(x, y + 24)
        else:
            pyautogui.doubleClick(x, y + 36)
    # Checks if "You need to change sites" message comes up
    time.sleep(1)
    pyautogui.click(900, 570)


def change_deposit_title():
    tour_menu = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
    while tour_menu is None:
        tour_menu = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                   region=(514, 245, 889, 566))
    x, y = tour_menu
    pyautogui.click(x + 300, y + 20)  # Tour Packages
    pyautogui.doubleClick(x + 300, y + 70)  # Top deposit in list
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changing_tour_package.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changing_tour_package.png',
            region=(514, 245, 889, 566))
    x, y = image
    pyautogui.click(x + 150, y + 125)  # Description
    keyboard.send('ctrl + z')  # Select all
    keyboard.send('ctrl + c')  # Copy description
    time.sleep(1)
    old_title = str(pyperclip.paste())
    new_title = old_title.replace("able", "ed")
    new_title = new_title.replace("ABLE", "ED")
    new_title = new_title.replace(" /", "/")
    new_title = new_title.replace("/ ", "/")
    keyboard.write(new_title)
    new_title = new_title.lower()

    if new_title == 'ir/refunded deposit':
        ir_refund(x, y, '50', '6/20/18')
    elif new_title == 'ams/refunded deposit':
        select_ams_refund_payment(x, y, '50', '6/20/18')


def copy_reference_number(x, y):
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


def select_ams_refund_payment(x, y, price, date):
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
    pyautogui.click(x + 300, y + 365)  # Payment button

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
        'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(147, 566, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(147, 566, 889, 566))
    x_2, y_2 = image

    pyautogui.click(x_2, y_2)  # AMS refund
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
    pyautogui.click(x - 20, y + 425)


def select_ams_refund_insert(x, y, price, date):
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
    pyautogui.click(x - 20, y + 425)


def select_ams_cc_ref(x, y, price, date):
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
    pyautogui.click(x - 20, y + 425)


def enter_rest_of_info(x, y, x_1, y_1):
    pyautogui.doubleClick(x_1 + 85, y + 130)  # Amount
    keyboard.write("50")
    pyautogui.doubleClick(x_1 + 85, y_1 + 140)  # Reference
    keyboard.send('ctrl + v')
    pyautogui.doubleClick(x_1 + 85, y_1 + 165)  # Date
    keyboard.write('6/21/18')
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
    pyautogui.click(x - 20, y + 425)


def ir_refund(x, y, price, date):
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
    pyautogui.click(x - 20, y + 425)


def error():
    pyautogui.click(x, y + 365)
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
        'C:\\Users\\Jared.Abrahams\\Screenshots\\void.png', region=(147, 566, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\void.png',
            region=(147, 566, 889, 566))
    x_2, y_2 = image
    pyautogui.click(x_2, y_2)  # Error
    pyautogui.doubleClick(x_1 + 85, y + 130)  # Amount
    keyboard.write("-50")
    pyautogui.doubleClick(x_1 + 85, y_1 + 140)  # Reference
    keyboard.write('ERROR')
    time.sleep(0.3)
    pyautogui.click(x_1 + 85, y_1 + 210)
    pyautogui.click(x_1 + 650, y_1 - 20)


pids = ['1388499', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '',
        '', '', '']
for pid in pids:
    search_pid(pid)
    select_tour()
    change_deposit_title()

'''search_pid('727085')
select_tour()
change_deposit_title()'''

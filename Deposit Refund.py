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

m1 = {}
m2 = {}
m3 = {}
m4 = {}
m5 = {}
m6 = {}
m7 = {}
m8 = {}


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
    m2['prospect_id'] = (m2_title[0] + 103, m2_title[1] + 50)
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
    m3['scroll_bar_wave'] = (m3_title[0] + 141, m3_title[1] + 384)
    #  Notes Tab
    m3['notes_change'] = (m3_title[0] + 50, m3_title[1] + 435)
    #  Tour Packages Tab
    m3['deposit_1'] = (m3_title[0] + 563, m3_title[1] + 71)
    m3['deposit_2'] = (m3_title[0] + 563, m3_title[1] + 94)
    m3['change_deposit'] = (m3_title[0] + 437, m3_title[1] + 181)
    #  Premiums Tab
    m3['premium_1'] = (m3_title[0] + 563, m3_title[1] + 64)
    m3['premium_2'] = (m3_title[0] + 563, m3_title[1] + 77)
    m3['premium_3'] = (m3_title[0] + 563, m3_title[1] + 90)
    m3['premium_4'] = (m3_title[0] + 563, m3_title[1] + 103)
    m3['premium_5'] = (m3_title[0] + 563, m3_title[1] + 116)
    m3['premium_6'] = (m3_title[0] + 563, m3_title[1] + 129)
    return m3


#  Select a Campaign
def get_m4_coordinates():
    global m4
    m4_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\titles\\select_a_campaign.png',
                                              region=(514, 245, 889, 566))
    while m4_title is None:
        m4_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\titles\\'
                                                  'select_a_campaign.png', region=(514, 245, 889, 566))

    m4['campaign'] = (m4_title[0] + 213, m4_title[1] + 118)
    m4['search'] = (m4_title[0] + 272, m4_title[1] + 118)
    m4['clear'] = (m4_title[0] + 356, m4_title[1] + 118)
    m4['first_campaign'] = (m4_title[0] + 214, m4_title[1] + 170)
    m4['select'] = (m4_title[0] + 251, m4_title[1] + 441)


# Co-prospect Menu
def get_m5_coordinates():
    global m5
    m5_title = pyautogui.locateCenterOnScreen(
        'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\record_will_be_added.png',
        region=(514, 245, 889, 566))
    while m5_title is None:
        m5_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles'
                                                  '\\record_will_be_added.png',
                                                  region=(514, 245, 889, 566))
    m5['get_from_prospect'] = (m5[0] + 198, m5[1] + 428)
    m5['first'] = (m5[0] + 212, m5[1] + 101)
    m5['ok'] = (m5[0] + 141, m5[1] + 462)


# Deposits menu
def get_m6_coordinates():
    global m6
    m6_title = pyautogui.locateCenterOnScreen(
        'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changing_tour_package.png', region=(514, 245, 889, 566))
    while m6_title is None:
        m6_title = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changing_tour_package.png',
            region=(514, 245, 889, 566))
    m6['title'] = (m6_title[0], m6_title[1])
    m6['description'] = (m6_title[0] + 172, m6_title[1] + 120)
    m6['deposit_1'] = (m6_title[0] + 294, m6_title[1] + 190)
    m6['insert'] = (m6_title[0], m6_title[1] + 369)
    m6['view'] = (m6_title[0] + 71, m6_title[1] + 370)
    m6['payment'] = (m6_title[0] + 300, m6_title[1] + 370)
    m6['ok'] = (m6_title[0] + 150, m6_title[1] + 406)
    m6['cancel'] = (m6_title[0] + 345, m6_title[1] + 406)


#  View Deposit
def get_m7_coordinates():
    global m7
    m7_title = pyautogui.locateCenterOnScreen(
        'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\view_record.png', region=(136, 652, 392, 247))
    while m7_title is None:
        m7_title = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\view_record.png', region=(136, 652, 392, 247))
    m7['title'] = (m7_title[0], m7_title[1])
    m7['transaction_code'] = (m7_title[0] + 223, m7_title[1] + 64)
    m7['description'] = (m7_title[0] + 273, m7_title[1] + 92)
    m7['amount'] = (m7_title[0] + 83, m7_title[1] + 117)
    m7['posted'] = (m7_title[0] + 253, m7_title[1] + 118)
    m7['reference'] = (m7_title[0] + 133, m7_title[1] + 144)
    m7['date'] = (m7_title[0] + 132, m7_title[1] + 170)
    m7['cancel'] = (m7_title[0] + 296, m7_title[1] + 215)


# Adding Deposit Menu
def get_m8_coordinates():
    global m8
    m8_title = pyautogui.locateCenterOnScreen(
        'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\adding_tour_package_item.png', region=(136, 652, 392, 247))
    while m8_title is None:
        m8_title = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\adding_tour_package_item.png', region=(136, 652, 392, 247))
    m8['title'] = (m8_title[0], m8_title[1])
    m8['transaction_code'] = (m8_title[0] + 179, m8_title[1] + 64)
    m8['description'] = (m8_title[0] + 229, m8_title[1] + 91)
    m8['amount'] = (m8_title[0] + 32, m8_title[1] + 117)
    m8['reference'] = (m8_title[0] + 87, m8_title[1] + 144)
    m8['date'] = (m8_title[0] + 87, m8_title[1] + 169)
    m8['ok'] = (m8_title[0] + 100, m8_title[1] + 214)
    m8['cancel'] = (m8_title[0] + 250, m8_title[1] + 215)


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


def check_tour_for_error():
    get_m3_coordinates()
    with mss.mss() as sct:
        x, y = m3['title']
        monitor = {'top': y + 171, 'left': x + 40, 'width': 52, 'height': 12}
        im = sct.grab(monitor)
        tour_status = str(mss.tools.to_png(im.rgb, im.size))
        if tour_status == sc.error:
            input('Is this the correct tour?')


def change_deposit_title(price):
    get_m3_coordinates()
    amount = 0
    x, y = m3['deposit_1']
    while amount != price:
        pyautogui.click(m3['tour_packages'])
        pyautogui.click(x, y)
        pyautogui.click(m3['change_deposit'])
        get_m6_coordinates()
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
        ir_refund(x, y, '50', '6/20/18')
    elif new_title.lower() == 'ams/refunded deposit':
        return "ams"
        select_ams_refund_payment(x, y, '50', '6/20/18')


"""def change_deposit_title():
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
    if new_title.lower() == 'ir/refunded deposit':
        ir_refund(x, y, '50', '6/20/18')
    elif new_title.lower() == 'ams/refunded deposit':
        select_ams_refund_payment(x, y, '50', '6/20/18')"""


def copy_reference_number():
    get_m6_coordinates()
    pyautogui.click(m6['view'])
    get_m7_coordinates()
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


def ams_cc_refund(date):
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


def select_ams_refund_payment(date, price, reference_number=None):
    get_m6_coordinates()
    pyautogui.click(m6['payment'])
    get_m8_coordinates()
    pyautogui.click(m8['transaction_code'])
    attempts = 0
    global transaction_code
    if transaction_code == 0 or transaction_code == 1:
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
    if transaction_code == 0 or transaction_code == 2:
        attempts = 0
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
    if transaction_code == 0 or transaction_code == 3:
        attempts = 0
        pyautogui.click(m8['cancel'])
        pyautogui.click(m6['insert'])
        get_m8_coordinates()
        pyautogui.click(m8['transaction_code'])
        image = pyautogui.locateCenterOnScreen(
            'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(136, 652, 392, 247))
        while image is None and attempts <= 2:
            image = pyautogui.locateCenterOnScreen(
                'C:\\Users\\Jared.Abrahams\\Screenshots\\ams_credit_refund.png', region=(136, 652, 392, 247))
        if image is not None:
            transaction_code = 3
        else:
            transaction_code = 0
            sys.exit("Couldn't find correct choice")
    pyautogui.click(image)
    if transaction_code == 3:
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


pids = ['', '1395277', '1397523', '1398810', '1406109', '1408288', '1408462', '1411866', '1413379', '1414706',
        '1416069', '1416722', '1418290', '1418362', '1418363', '1392622', '1395491', '1395564', '1397310', '1400120',
        '1400121', '1405961', '1408413', '1408761', '1412281', '1415279', '1417083', '1418375', '1418434', '1418441',
        '1418479', '1418487', '1418501', '1418504', '1418512', '1418535', '1392675', '1394122', '1397433', '1401109',
        '1408282', '1412152', '1413928', '1413964', '1413976', '1414268', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '']
transaction_code = 0
date = input("Date")
price = input("Price")
for pid in pids:
    if pid != '':
        search_pid(pid)
        select_tour()
        check_tour_for_error()
        deposit_type = change_deposit_title(price)
        if deposit_type != 'prev':
            copy_reference_number()
        else:
            ams_or_ir = input('ams or ir:')
            reference_number = input('Reference Number starting with R:')
            if ams_or_ir == 'ams':
                select_ams_refund_payment(date, price, reference_number)
        if deposit_type == 'ams':
            select_ams_refund_payment(date, price)

'''search_pid('727085')
select_tour()
change_deposit_title()'''

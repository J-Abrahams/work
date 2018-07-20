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

m1 = {}
m2 = {}
m3 = {}


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


def check_tour_for_error():
    sc.get_m3_coordinates()
    with mss.mss() as sct:
        x, y = m3['title']
        monitor = {'top': y + 171, 'left': x + 40, 'width': 52, 'height': 12}
        im = sct.grab(monitor)
        tour_status = str(mss.tools.to_png(im.rgb, im.size))
        if tour_status == sc.error:
            input('Is this the correct tour?')


def check_for_refundable_deposit():
    pyautogui.click(m3['tour_packages'])
    x, y = m3['deposit_1']
    is_deposit_blue = pyautogui.pixelMatchesColor(x, y, (8, 36, 107))
    while is_deposit_blue is True:
        pyautogui.click(m3['change_deposit'])
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
            pyautogui.click(x_2 + 85, y_2 + 370)
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


def check_for_dep_premium():
    sc.get_m3_coordinates()
    price = check_for_refundable_deposit()
    premiums = read_premiums()
    if price == '40':
        if any(sc.dep_40_cc in s for s in premiums) or any(sc.dep_40_cash in s for s in premiums) or \
                any(sc.d40_cc_dep in s for s in premiums) or any(sc.d40_dep in s for s in premiums):
            print('\x1b[6;30;42m' + '$40 DEP is present' + '\x1b[0m')
        else:
            print('\x1b[6;30;41m' + 'Missing $40 DEP' + '\x1b[0m')
    elif price == '50':
        if any(sc.dep_50_cc in s for s in premiums) or any(sc.dep_50_cash in s for s in premiums) or \
                any(sc.d50_cc_dep in s for s in premiums):
            print('\x1b[6;30;42m' + '$50 DEP is present' + '\x1b[0m')
        else:
            print('\x1b[6;30;41m' + 'Missing $50 DEP' + '\x1b[0m')
    elif price == '20':
        if any(sc.d20_cc_dep in s for s in premiums) or any(sc.dep_20_cc in s for s in premiums):
            print('\x1b[6;30;42m' + '$20 DEP is present' + '\x1b[0m')
        else:
            print('\x1b[6;30;41m' + 'Missing $20 DEP' + '\x1b[0m')


def read_premiums():
    """Adds all premiums to the list premiums and checks if there are any duplicates among them"""
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
    else:
        print('\x1b[6;30;42m' + str(number_of_premiums) + ' Premiums - No Duplicates' + '\x1b[0m')
    return list_of_premiums


def confirm_tour_type():
    pass


def confirm_tour_status(status):
    """Checks that the tour status is correct"""
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
    if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_confirmer_sol.png',
                                      region=(514, 245, 889, 566)) is None:
        print('\x1b[6;30;42m' + 'Sol number is good' + '\x1b[0m')
    else:
        pyautogui.doubleClick(x + 115, y + 222)
        keyboard.press_and_release('ctrl + c')
        r = Tk()
        tsw_sol = str(r.selection_get(selection="CLIPBOARD"))
        print(tsw_sol + " changed to " + sol)
        pyperclip.copy(sol)
        keyboard.press_and_release('ctrl + v')
    pyautogui.click(x - 65, y + 18)


def notes(status):
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
                print('\x1b[6;30;41m' + 'COULDN\'T FIND CORRECT NOTE' + '\x1b[0m')
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
                                                      region=(514, 245, 889, 566)) is None:
                        print('\x1b[6;30;42m' + 'Tour Result is correct' + '\x1b[0m')
                    else:
                        print('\x1b[6;30;41m' + 'NO TOUR RESULT' + '\x1b[0m')
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
            return


def enter_personnel(sol, status):
    sc.get_m3_coordinates()
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
    if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_confirmer.png',
                                      region=(514, 245, 889, 566)) is None:
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


def manual_confirmation(pids):
    status = input('Conf (c), RXL (r), CXL (x), UG (u), or TAV (tav)')
    for pid in pids:
        if pid != '':
            search_pid(pid)
            double_check_pid(pid)
            select_tour()
            check_tour_for_error()
            notes(status)
            check_for_dep_premium()
            confirm_tour_status(status)
            enter_personnel(sol, status)
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
            conf = row['conf']
            cxl = row['cxl']
            rxl = row['rxl']
            search_pid(pids)
            double_check_pid(pids)
            select_tour()
            check_tour_for_error()
            check_for_dep_premium()
            try:
                ug = row['ug']
                if ug == "X" or ug == "x":
                    enter_personnel(sol, 'u')
            except KeyError:
                pass
            try:
                tav = row['tav']
                if tav == "X" or tav == "x":
                    enter_personnel(sol, 'tav')
            except KeyError:
                pass
            if conf == "X" or conf == "x":
                confirm_tour_status('c')
                notes('c')
                enter_personnel(sol, 'c')
            if rxl == "X" or rxl == "x":
                confirm_tour_status('r')
                notes('r')
                enter_personnel(sol, 'r')
            if cxl == "X" or cxl == "x":
                confirm_tour_status('x')
                notes('x')
                enter_personnel(sol, 'x')

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


'''df = pd.read_csv("file.csv", header=None, skiprows=3)
for row in df:
    pids = row['PID']
    print(pids)
    pids = pids[:-2]
    conf = row['CONF']
    cxl = row['CXL']
    rxl = row['RXL']
    search_pid(pids)
    select_tour()
    add_personnel('sol25688', conf, cxl, rxl)

df.SP.head(2)'''

pids = ['1326419', '1412106', '1419962', '1412393', '1417981', '1418751', '1418279', '1417363', '', '',
        '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '']

auto_or_manual = input('Auto (A) or Manual (M):')
print("Justin Locke - 4967 \nBrian Bennett - 3055")
sol = "SOL" + input("SOL #:")
if auto_or_manual == 'a' or auto_or_manual == 'A':
    automatic_confirmation()
else:
    manual_confirmation(pids)

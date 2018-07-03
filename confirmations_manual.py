import keyboard
import pyautogui
import time
import pyperclip
import csv


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
    m3['premium_1'] = (m3_title[0] + 333, m3_title[1] + 64)
    m3['premium_2'] = (m3_title[0] + 333, m3_title[1] + 78)
    m3['premium_3'] = (m3_title[0] + 333, m3_title[1] + 90)
    m3['premium_4'] = (m3_title[0] + 333, m3_title[1] + 102)
    m3['premium_5'] = (m3_title[0] + 333, m3_title[1] + 114)
    m3['premium_6'] = (m3_title[0] + 333, m3_title[1] + 126)
    return m3


def count_number_of_premiums():
    m3 = get_m3_coordinates()
    pyautogui.click(m3['premium_6'])
    x, y = m3['premium_6']
    if pyautogui.pixelMatchesColor(x, y, (8, 36, 107)) is True:
        print("6 Premiums")
        return
    else:
        pyautogui.click(m3['premium_5'])
    x, y = m3['premium_5']
    if pyautogui.pixelMatchesColor(x, y, (8, 36, 107)) is True:
        print("5 Premiums")
        return
    else:
        pyautogui.click(m3['premium_4'])
    x, y = m3['premium_4']
    if pyautogui.pixelMatchesColor(x, y, (8, 36, 107)) is True:
        print("4 Premiums")
        return
    else:
        pyautogui.click(m3['premium_3'])
    x, y = m3['premium_3']
    if pyautogui.pixelMatchesColor(x, y, (8, 36, 107)) is True:
        print("3 Premiums")
        return
    else:
        pyautogui.click(m3['premium_2'])
    x, y = m3['premium_2']
    if pyautogui.pixelMatchesColor(x, y, (8, 36, 107)) is True:
        print("2 Premiums")
        return
    else:
        pyautogui.click(m3['premium_1'])
    x, y = m3['premium_1']
    if pyautogui.pixelMatchesColor(x, y, (8, 36, 107)) is True:
        print("1 Premiums")
        return
    else:
        print("No Premiums")


def search_pid(pid_number):
    m1 = get_m1_coordinates()
    pyautogui.doubleClick(m1['search'])
    keyboard.write(pid_number)
    pyautogui.click(m1['find_now'])
    pyautogui.click(m1['change'])


def select_tour():
    m2 = get_m2_coordinates()
    audition = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_audition.png',
                                              region=(514, 245, 889, 566))
    if audition is None:
        pyautogui.doubleClick(m2['first_tour'])

    else:
        audition = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\audition_2.png',
                                                  region=(514, 245, 889, 566))
        if audition is None:
            pyautogui.doubleClick(m2['second_tour'])
        else:
            pyautogui.doubleClick(m2['third_tour'])
    # Checks if "You need to change sites" message comes up
    time.sleep(1)
    pyautogui.click(m2['yes_change_sites'])


def confirm_tour_status_confirm():
    x, y = m3['title']
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
        print("TOUR STATUS MIGHT BE INCORRECT")
    else:
        print("Tour status is good")


def confirm_tour_status_reschedule():
    x, y = m3['title']
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\rescheduled.png',
                                           region=(x + 27, y + 132, 131, 103))
    attempts = 0
    while image is None and attempts <= 2:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\rescheduled.png',
                                               region=(x + 27, y + 132, 131, 103))
        attempts += 1

    if image is None:
        print("TOUR STATUS MIGHT BE INCORRECT")
    else:
        print("Tour status is good")


def confirm_tour_status_cancel():
    x, y = m3['title']
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\canceled.png',
                                           region=(x + 27, y + 132, 131, 103))
    attempts = 0
    while image is None and attempts <= 2:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\canceled.png',
                                               region=(x + 27, y + 132, 131, 103))
        attempts = + 1

    if image is None:
        print("TOUR STATUS MIGHT BE INCORRECT")
    else:
        print("Tour status is good")


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
        time.sleep(1)
        print(tsw_sol + " changed to " + sol)
        pyperclip.copy(sol)
        keyboard.press_and_release('ctrl + v')

    except TypeError:
        print("Sol number is good")

    pyautogui.click(x - 65, y + 18)


def check_deposit():
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    pyautogui.click(x + 300, y + 18)


def confirm(sol):
    get_m3_coordinates()
    confirm_tour_status_confirm()
    confirm_sol_in_userfields(sol)
    pyautogui.click(m3['personnel'])
    pyautogui.click(m3['insert'])
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                               region=(514, 245, 889, 566))
    x_1, y_1 = image
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
    except Exception:
        pyautogui.click(x_4 + 90, y_4 + 80)
        keyboard.write("cc")
    pyautogui.click(x_4 + 90, y_4 + 105)
    keyboard.write("cc")
    pyautogui.click(x_4 + 90, y_4 + 350)


def reschedule(sol):
    get_m3_coordinates()
    confirm_tour_status_reschedule()
    pyautogui.click(m3['personnel'])
    pyautogui.click(m3['insert'])
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                               region=(514, 245, 889, 566))
    x_1, y_1 = image
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
    except Exception:
        pyautogui.click(x_4 + 90, y_4 + 80)
        keyboard.write("cc")
    pyautogui.click(x_4 + 90, y_4 + 105)
    keyboard.write("r")
    pyautogui.click(x_4 + 90, y_4 + 350)


def cancel(sol):
    get_m3_coordinates()
    confirm_tour_status_cancel()
    pyautogui.click(m3['personnel'])
    pyautogui.click(m3['insert'])
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                               region=(514, 245, 889, 566))
    x_1, y_1 = image
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
    except Exception:
        pyautogui.click(x_4 + 90, y_4 + 80)
        keyboard.write("cc")
    pyautogui.click(x_4 + 90, y_4 + 105)
    keyboard.write("c")
    pyautogui.click(x_4 + 90, y_4 + 350)


def upgrade(sol):
    get_m3_coordinates()
    pyautogui.click(m3['personnel'])
    pyautogui.click(m3['insert'])
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                               region=(514, 245, 889, 566))
    x_1, y_1 = image
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
    except Exception:
        pyautogui.click(x_4 + 90, y_4 + 80)
        keyboard.write("cc")
    pyautogui.click(x_4 + 90, y_4 + 105)
    keyboard.write("u")
    pyautogui.click(x_4 + 90, y_4 + 350)


def travel_allowance(sol):
    get_m3_coordinates()
    pyautogui.click(m3['personnel'])
    pyautogui.click(m3['insert'])
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\t_personnel.png',
                                               region=(514, 245, 889, 566))
    x_1, y_1 = image
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
    except Exception:
        pyautogui.click(x_4 + 90, y_4 + 80)
        keyboard.write("cc")
    pyautogui.click(x_4 + 90, y_4 + 105)
    keyboard.write("t")
    pyautogui.click(x_4 + 90, y_4 + 350)


# c = confirm, x = cancel, r = reschedule, t = TAV, u = upgrade
sol_num = "sol2956"
status = 'r'
pids = ['1411763', '1331282', '1415096', '', '', '', '', '', '', '',
        '', '', '', '', '', '', '', '', '', '',
        '', '', '', '', '', '']
for pid in pids:
    if pid != '':
        search_pid(pid)
        select_tour()
        if status == "c":
            confirm(sol_num)
        elif status == "r":
            reschedule(sol_num)
        elif status == "x":
            cancel(sol_num)
        elif status == "u":
            upgrade(sol_num)
        elif status == "t":
            travel_allowance(sol_num)
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

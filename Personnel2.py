import keyboard
import pyautogui
import time
import pyperclip
import csv


def search_pid(pid):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png',
                                           region=(514, 245, 889, 566))
    while image == None:
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
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    audition = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_audition.png',
                                              region=(514, 245, 889, 566))
    if audition == None:
        pyautogui.doubleClick(x, y + 12)

    else:
        pyautogui.doubleClick(x, y + 24)
    # Checks if "You need to change sites" message comes up
    time.sleep(1)
    pyautogui.click(900, 570)
    """img = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date2.png')
    if img == None:
        x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_yes.png')
        pyautogui.doubleClick(x, y)
        time.sleep(1)"""


"""x, y = pyautogui.locateCenterOnScreen("minivac_showed_sh.png")
pyautogui.doubleClick(x, y)
pyautogui.click(x=900, y=560)

time.sleep(3)"""


def confirm_sol_in_userfields(sol):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                           region=(514, 245, 889, 566))
    while image == None:
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
        print("Sol number was good")
    
    pyautogui.click(x - 65, y + 18)


def check_deposit():
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                           region=(514, 245, 889, 566))
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    pyautogui.click(x + 300, y + 18)

def add_personnel(sol, conf, cxl, rxl):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                          region=(514, 245, 889, 566))
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    pyautogui.click(x + 200, y + 220) #Personnel Tab
    pyautogui.click(x + 275, y + 440) #Insert Button
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_personnel_menu.png',
                                           region=(514, 245, 889, 566))
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_personnel_menu.png',
                                               region=(514, 245, 889, 566))
    x_1, y_1 = image
    pyautogui.click(x_1 + 75, y_1 + 25) #By Personnel Number Tab
    keyboard.write(sol)
    pyautogui.doubleClick(x_1, y_1 + 100) #Person in list
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_personnel_titles_menu.png',
                                           region=(514, 245, 889, 566))
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_personnel_titles_menu.png',
                                               region=(514, 245, 889, 566))
    x_3, y_3 = image
    pyautogui.click(x_3 + 75, y_3 + 150) #Close
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_adding_records_menu.png',
                                           region=(514, 245, 889, 566))
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_adding_records_menu.png',
                                               region=(514, 245, 889, 566))
    x_4, y_4 = image
    try: 
        x_5, y_5 = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_confirmer.png', region=(514, 245, 889, 566))
    except Exception:
        pyautogui.click(x_4 + 90, y_4 + 80)
        keyboard.write("cc")
    pyautogui.click(x_4 + 90, y_4 + 105)
    
    if conf == "X":
        keyboard.write("cc")

    if cxl == "X":
        keyboard.write("c")
    
    if rxl == "X":
        keyboard.write("r")
        
    """elif status == "t":
        keyboard.write("t")
        
    elif status == "u":
        keyboard.write("u")

    elif status == "tav":
        keyboard.write("t")"""
    
    pyautogui.click(x_4 + 90, y_4 + 350)

    keep_going = input("Everything ok?")

    if keep_going != '':
        pyautogui.click(x + 265, y + 475)
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                               region=(514, 245, 889, 566))
        while image == None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                                   region=(514, 245, 889, 566))
        x, y = image
        pyautogui.click(x - 20, y + 425)


#c = confirm, x = cancel, r = reschedule, t = TAV, u = upgrade
"""id_num = "1412216"
sol_num, status = "sol23521", 'x'
search_pid(id_num)
select_tour()
if status == "c":
    confirm_sol_in_userfields(sol_num)
add_personnel(sol_num, status)
#add_personnel(sol_num, 'u')
#welk_workers = {Katherine_Albini:SOL23521}"""
"""sol_num = 'sol23542'
pids = [['1412478', 'x'], ['1413013', 'c'], ['1412008', 'r'], ['1410649', 'x'], ['1413089', 'r'], ['1412926', 'c'],
        ['1410674', 'c'], ['1410035', 'u']]  # ['1412854', 'c'], ['1397766', 'r'], ['1409737', 'c'], ['1402431', 'r']]
for pid, status in pids:
    print(pid + status)
    search_pid(pid)
    select_tour()
    if status == "c":
        confirm_sol_in_userfields(sol_num)
    add_personnel(sol_num, status)"""

with open('file.csv') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        pids = row['PID']
        pids = pids[:-2]
        conf = row['Conf']
        cxl = row['Cxl']
        rxl = row['Fix!']
        search_pid(pids)
        select_tour()
        add_personnel('sol23521', conf, cxl, rxl)
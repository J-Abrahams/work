import keyboard
import pyautogui
import time
import pyperclip


def search_pid(pid):
    x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png')
    pyautogui.doubleClick(x + 50, y)
    keyboard.write(id)
    pyautogui.click(x + 650, y)
    pyautogui.click(x + 400, y + 500)
    time.sleep(4)

def select_tour():
    x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png')
    pyautogui.doubleClick(x, y + 12)
    #Checks if "You need to change sites" message comes up
    time.sleep(1)
    img = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date2.png')
    if img == None:
        x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_yes.png')
        pyautogui.doubleClick(x, y)
        time.sleep(1)

"""x, y = pyautogui.locateCenterOnScreen("minivac_showed_sh.png")
pyautogui.doubleClick(x, y)
pyautogui.click(x=900, y=560)

time.sleep(3)"""


def add_personnel(sol, status):
    x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png', region=(514, 245, 889, 566))
    pyautogui.click(x + 200, y + 220) #Personnel Tab
    pyautogui.click(x + 275, y + 440) #Insert Button
    time.sleep(0.75)
    x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_personnel_menu.png', region=(514, 245, 889, 566))
    pyautogui.click(x + 75, y + 25) #By Personnel Number Tab
    keyboard.write(sol)
    pyautogui.doubleClick(x, y + 100) #Person in list
    time.sleep(0.5)
    x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_personnel_titles_menu.png', region=(514, 245, 889, 566))
    pyautogui.click(x + 75, y + 150) #Close
    x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_adding_records_menu.png', region=(514, 245, 889, 566))
    pyautogui.click(x + 90, y + 80)
    keyboard.write("q")
    if status != "q":
        pyautogui.click(x + 90, y + 105)
        
        if status == "x":
            keyboard.write("c")

        elif status == "r":
            keyboard.write("r")
    
        elif status == "u":
            keyboard.write("u")

    pyautogui.click(x + 90, y + 350)


#c = confirm, x = cancel, r = reschedule, t = TAV, u = upgrade
id_num = "517263"
sol_num, status = "SOL25699", "x"

add_personnel(sol_num, status)

#welk_workers = {Katherine_Albini:SOL23521}

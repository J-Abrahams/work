def search_pid(pid):
    x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png', region=(514, 245, 889, 566))
    pyautogui.doubleClick(x + 50, y)
    keyboard.write(id)
    pyautogui.click(x + 650, y)
    pyautogui.click(x + 400, y + 500)
    #time.sleep(4)
    
def select_tour():
    try:
        x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png', region=(514, 245, 889, 566))
    
    except Exception:
        return select_tour()
    
    if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_audition.png') != None:
        x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_audition.png', region=(945, 304, 190, 86))
        pyautogui.doubleClick(x, y + 20)
    else:
        x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png', region=(514, 245, 889, 566))
        pyautogui.doubleClick(x, y + 12)  #Checks if "You need to change sites" message comes up
    time.sleep(2)
    img = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date2.png', region=(514, 245, 889, 566))
    if img == None:
        x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_yes.png', region=(514, 245, 889, 566))
        pyautogui.doubleClick(x, y)
        time.sleep(1)
        
def create_dep():
    x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png', region=(514, 245, 889, 566))
    pyautogui.click(x + 300, y + 20)
    try:
        x, y = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_deposit.png', region=(514, 245, 889, 566))
    
    except Exception:
        amount = input('Deposit amount:)
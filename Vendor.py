import keyboard
import pyautogui
import time
import pyperclip

def switch_site(site):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\windows_closed.png',
                                       region=(514, 245, 300, 300))
    if image == None:
        print("Close all windows")
        raise SystemExit(0)

    pyautogui.click(10, 25)
    pyautogui.click(10, 150)
    pyautogui.click(1000, 375)
    if data_dict['location'] == '2':
        keyboard.write('A1')
        keyboard.send('enter')
    elif data_dict['location'] == '3':
        keyboard.write('A2')
        keyboard.send('enter')
    elif data_dict['location'] == '4':
        keyboard.write('A3')
        keyboard.send('enter')
    elif data_dict['location'] == '5':
        keyboard.write('T')
        keyboard.send('enter')
    elif data_dict['location'] == '8':
        keyboard.write('C')
        keyboard.send('enter')
    elif data_dict['location'] == '9':
        keyboard.write('Welk Resort N')
        keyboard.send('enter')
    elif data_dict['location'] == '11':
        keyboard.write('Welk Resort Bre')
        keyboard.send('enter')
    pyautogui.click(10, 25)
    pyautogui.click(10, 45)


def search_pid(pid):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png',
                                           region=(514, 245, 889, 566))
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    pyautogui.doubleClick(x + 50, y)
    keyboard.write(data_dict['pid'])
    pyautogui.click(x + 650, y)
    pyautogui.click(x + 400, y + 500)


def enter_prospect_info(info):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changingaprospect.png',
                                           region=(514, 245, 889, 566))
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\changingaprospect.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    pyautogui.doubleClick(x + 135, y + 85)
    for i in range(35):
        keyboard.send('backspace')
    keyboard.write(data_dict['last name'])
    pyautogui.doubleClick(x + 280, y + 85)
    for i in range(35):
        keyboard.send('backspace')
    keyboard.write(data_dict['first_name'])
    pyautogui.doubleClick(x + 135, y + 300)
    time.sleep(0.3)
    keyboard.press_and_release('ctrl + c')
    time.sleep(0.3)
    copied_number_1 = pyperclip.paste()
    pyautogui.doubleClick(x + 280, y + 300)
    time.sleep(0.3)
    keyboard.press_and_release('ctrl + c')
    copied_number_2 = pyperclip.paste()
    print(data_dict['home_phone'])
    print(copied_number_1)
    print(copied_number_2)
    if data_dict['home_phone'] != copied_number_1 and copied_number_2:
        pyautogui.doubleClick(x + 135, y + 300)
        keyboard.write(data_dict['home_phone'])
    if data_dict['email'] != '':
        pyautogui.doubleClick(x + 280, y + 350)
    pyautogui.doubleClick(x + 60, y + 25)
    if data_dict['spouse_first_name'] and data_dict['spouse_last_name'] != '':
        pyautogui.doubleClick(x + 280, y + 115)
        for i in range(35):
            keyboard.send('backspace')
        keyboard.write(data_dict['spouse_first_name'] + ' ' + data_dict['spouse_last_name'])
    pyautogui.doubleClick(x + 225, y + 380)
    keyboard.write(data_dict['credit_card_#'])
    pyautogui.doubleClick(x + 135, y + 425)
    keyboard.write(data_dict['expiration_date'])
    keyboard.send('tab')
    pyautogui.click(x + 480, y + 180)
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                           region=(514, 245, 889, 566))
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
    x_1, y_1 = image
    pyautogui.click(x_1 + 155, y_1 + 125)
    pyautogui.click(x_1 - 125, y_1 + 5)
    keyboard.write(data_dict['source_code'])
    keyboard.send('enter')
    pyautogui.click(x_1 - 125, y_1 + 60)
    pyautogui.click(x_1 - 60, y_1 + 325)
    pyautogui.click(x_1 + 130, y_1 + 150)
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\minivac.png',
                                           region=(514, 245, 889, 566))
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\minivac.png',
                                               region=(514, 245, 889, 566))
    x_2, y_2 = image
    pyautogui.click(x_2, y_2)
    pyautogui.click(x_1 + 130, y_1 + 175)
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\booked.png',
                                           region=(514, 245, 889, 566))
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\booked.png',
                                               region=(514, 245, 889, 566))
    x_2, y_2 = image
    pyautogui.click(x_2, y_2)
    pyautogui.click(x_1 + 130, y_1 + 200)
    keyboard.write(data_dict['tour_date'])
    pyautogui.click(x_1 + 325, y_1 + 180)

    '''time.sleep(1)
    keyboard.write('m')
    keyboard.send('tab')
    keyboard.write('b')
    keyboard.send('tab')
    keyboard.write(data_dict['tour_date'])
    keyboard.send('tab')
    for i in range(10):
        keyboard.send('down')'''

data_dict = {
    'agent_name': 'Borges',
    'location': '2',
    'source_code': 'OMOWNMM',
    'pid': '589850',
    'first_name': 'Stanley D.',
    'last name': 'Prause',
    'spouse_first_name': 'Nancy E.',
    'spouse_last_name': 'Prause',
    'home_phone': '951-791-1931',
    'other_phone': '',
    'email': 'watanga@gmail.com',
    'tour_date': '8/5',
    'tour_time': '1300',
    'type_of_deposit': 'refundable',
    'deposit_amount': '99',
    'expiration_date': '8/19',
    'cvv_code': '818',
    'credit_card_#': '4815881027267355',
    'accommodation': 'vvl',
    'arrival_date': '8/5',
    'number_of_nights': '2',
    'number_of_adults': '2',
    'number_of_kids': '0'
}


switch_site(data_dict)
search_pid(data_dict)
enter_prospect_info(data_dict)


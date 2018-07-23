import keyboard
import pyautogui
import time
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8


def switch_site(site):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\windows_closed.png',
                                           region=(514, 245, 300, 300))
    if image is None:
        print("Close all windows")
        raise SystemExit(0)

    pyautogui.click(10, 25)
    pyautogui.click(10, 150)
    pyautogui.click(1000, 375)
    if owner_dict['location'] == '2':
        keyboard.write('A1')
        keyboard.send('enter')
    elif owner_dict['location'] == '3':
        keyboard.write('A2')
        keyboard.send('enter')
    elif owner_dict['location'] == '4':
        keyboard.write('A3')
        keyboard.send('enter')
    elif owner_dict['location'] == '5':
        keyboard.write('T')
        keyboard.send('enter')
    elif owner_dict['location'] == '8':
        keyboard.write('C')
        keyboard.send('enter')
    elif owner_dict['location'] == '9':
        keyboard.write('Welk Resort N')
        keyboard.send('enter')
    elif owner_dict['location'] == '11':
        keyboard.write('Welk Resort Bre')
        keyboard.send('enter')
    pyautogui.click(10, 25)
    pyautogui.click(10, 45)


def search_pid(pid_number):
    sc.get_m1_coordinates()
    pyautogui.doubleClick(m1['search'])
    keyboard.write(owner_dict['pid'])
    pyautogui.click(m1['find_now'])
    pyautogui.click(m1['change'])


def enter_card(owner_dict):
    sc.get_m2_coordinates()
    pyautogui.click(m2['demographics'])
    pyautogui.doubleClick(m2['card_number'])
    keyboard.write(owner_dict['card_number'])
    pyautogui.doubleClick(m2['expiration'])
    keyboard.write(owner_dict['expiration'])
    pyautogui.click(m2['insert_tour'])


def enter_tour_info(owner_dict):
    sc.get_m3_coordinates()
    pyautogui.click(m3['campaign'])
    #  Menu 4 - Select a Campaign
    sc.get_m4_coordinates()
    pyautogui.click(m4['clear'])
    pyautogui.click(m4['campaign'])
    keyboard.write(owner_dict['campaign'])
    pyautogui.click(m4['select'])
    # Menu 3 - Adding a Tour Record
    sc.get_m3_coordinates()
    pyautogui.click(m3['tour_type'])
    keyboard.write(owner_dict['tour_type'])
    pyautogui.click(m3['tour_status'])
    keyboard.write('b')
    pyautogui.click(m3['tour_date'])
    keyboard.write(owner_dict['tour_date'])
    pyautogui.click(m3['tour_location'])
    for i in range(5):
        keyboard.send('down')
    pyautogui.click(m3['title'])
    pyautogui.click(m3['wave'])
    if owner_dict['tour_time'] == "800":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_800.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "815":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_815.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "830":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_830.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "900":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_900.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "915":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_915.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "930":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_930.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "1030":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1030.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "1045":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1045.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "1130":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1130.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "1145":
        pyautogui.click(m3['scroll_bar_wave'])
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1145.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "1230":
        pyautogui.click(m3['scroll_bar_wave'])
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1230.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "1300":
        pyautogui.click(m3['scroll_bar_wave'])
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1300.png',
                                                       region=(514, 245, 889, 566)))
    elif owner_dict['tour_time'] == "1315":
        pyautogui.click(m3['scroll_bar_wave'])
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1315.png',
                                                       region=(514, 245, 889, 566)))


def create_accommodation(owner_dict):
    pyautogui.click(m3['insert'])


owner_dict = {
    'agent_name': 'Haley',
    #  BR = 5, DO = 3, LT = 9, SD = 3
    'location': '2',
    'campaign': 'OMOWNMM',
    'pid': '241467',
    'first_name': 'Bassam',
    'last name': 'Jaradat',
    'tour_date': '8/27',
    'tour_time': '1230',
    'type_of_deposit': 'Refundable',
    'deposit_amount': '99',
    'card_number': '440066703897',
    'expiration': '01/20',
    'cvv_code': '846',
    'accommodation': 'vvl',
    'arrival_date': '7/7',
    'number_of_nights': '3',
    'number_of_adults': '2',
    'number_of_kids': ''
}


switch_site(owner_dict)
search_pid(owner_dict)
enter_card(owner_dict)
enter_tour_info(owner_dict)

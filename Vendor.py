import keyboard
import pyautogui
import screenshot_data as sc

m1 = {}
m2 = {}
m3 = {}
m4 = {}
m5 = {}


def switch_site(vendor_dict):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\windows_closed.png',
                                           region=(514, 245, 300, 300))
    if image is None:
        print("Close all windows")
        raise SystemExit(0)

    pyautogui.click(10, 25)
    pyautogui.click(10, 150)
    pyautogui.click(1000, 375)
    if vendor_dict['location'] == '2':
        keyboard.write('A1')
        keyboard.send('enter')
    elif vendor_dict['location'] == '3':
        keyboard.write('A2')
        keyboard.send('enter')
    elif vendor_dict['location'] == '4':
        keyboard.write('A3')
        keyboard.send('enter')
    elif vendor_dict['location'] == '5':
        keyboard.write('T')
        keyboard.send('enter')
    elif vendor_dict['location'] == '8':
        keyboard.write('C')
        keyboard.send('enter')
    elif vendor_dict['location'] == '9':
        keyboard.write('Welk Resort N')
        keyboard.send('enter')
    elif vendor_dict['location'] == '11':
        keyboard.write('Welk Resort Bre')
        keyboard.send('enter')
    pyautogui.click(10, 25)
    pyautogui.click(10, 45)


def insert_new_pid():
    m1 = sc.get_m1_coordinates()
    pyautogui.click(m1['insert'])


def enter_prospect_info(vendor_dict):
    m2 = sc.get_m2_coordinates()
    pyautogui.click(m2['last_name'])
    keyboard.write(vendor_dict['last name'])
    pyautogui.click(m2['first_name'])
    keyboard.write(vendor_dict['first_name'])
    pyautogui.click(m2['address'])
    keyboard.write(vendor_dict['address'])
    pyautogui.click(m2['city'])
    keyboard.write(vendor_dict['city'])
    pyautogui.click(m2['state'])

    pyautogui.click(m2['postal_code'])
    keyboard.write(vendor_dict['postal_code'])
    pyautogui.click(m2['country'])
    for i in range(5):
        keyboard.send('u')
    pyautogui.click(m2['phone1'])
    keyboard.write(vendor_dict['phone1'])
    if vendor_dict['phone2'] != '':
        pyautogui.click(m2['phone2'])
        keyboard.write(vendor_dict['phone2'])
    if vendor_dict['email'] != '':
        pyautogui.click(m2['email'])
        keyboard.write(vendor_dict['email'])
    pyautogui.click(m2['demographics'])
    #  Demographics Tab
    pyautogui.click(m2['marital_status'])
    if vendor_dict['marital_status'] == 'm':
        keyboard.send('m')
    pyautogui.click(m2['spouse'])
    keyboard.write(vendor_dict['spouse_first_name'] + ' ' + vendor_dict['spouse_last_name'])
    pyautogui.click(m2['occupation'])
    keyboard.send('e')
    pyautogui.click(m2['income'])
    keyboard.write(vendor_dict['income'])
    pyautogui.click(m2['preferred_language'])
    keyboard.send('e')
    # pyautogui.doubleClick(m2['card_number'])
    # keyboard.write(vendor_dict['card_number'])
    # pyautogui.doubleClick(m2['expiration'])
    # keyboard.write(vendor_dict['expiration'])
    pyautogui.click(m2['notes_co_tab'])
    pyautogui.click(m2['insert_coprospect'])
    # Menu 4 - Adding a co-prospect
    sc.get_m5_coordinates()
    pyautogui.click(m5['get_from_prospect'])
    pyautogui.click(m5['first'])
    keyboard.send(vendor_dict['spouse_first_name'])
    pyautogui.click(m5['ok'])
    pyautogui.click(m2['insert_tour'])
    #  Menu 3 - Adding a Tour Record
    sc.get_m3_coordinates()
    pyautogui.click(m3['campaign'])
    #  Menu 4 - Select a Campaign
    sc.get_m4_coordinates()
    pyautogui.click(m4['clear'])
    pyautogui.click(m4['campaign'])
    keyboard.write(vendor_dict['campaign'])
    pyautogui.click(m4['select'])
    # Menu 3 - Adding a Tour Record
    sc.get_m3_coordinates()
    pyautogui.click(m3['tour_type'])
    keyboard.write(vendor_dict['tour_type'])
    pyautogui.click(m3['tour_status'])
    keyboard.write('b')
    pyautogui.click(m3['tour_date'])
    keyboard.write(vendor_dict['tour_date'])
    pyautogui.click(m3['tour_location'])
    for i in range(5):
        keyboard.send('down')
    pyautogui.click(m3['wave'])
    if vendor_dict['tour_time'] == "800":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_800.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "815":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_815.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "830":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_830.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "900":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_900.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "915":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_915.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "930":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_930.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "1030":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1030.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "1045":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1045.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "1130":
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1130.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "1145":
        pyautogui.click(m3['scroll_bar_wave'])
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1145.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "1230":
        pyautogui.click(m3['scroll_bar_wave'])
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1230.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "1300":
        pyautogui.click(m3['scroll_bar_wave'])
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1300.png',
                                                       region=(514, 245, 889, 566)))
    elif vendor_dict['tour_time'] == "1315":
        pyautogui.click(m3['scroll_bar_wave'])
        pyautogui.click(pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1315.png',
                                                       region=(514, 245, 889, 566)))


data_dict = {
    'agent_name': '',
    'location': 'br',
    'campaign': 'OMOWNMM',
    'pid': '',
    'first_name': 'Bassam',
    'last name': 'Jaradat',
    'spouse_first_name': 'Fatheil',
    'spouse_last_name': 'Abdallah',
    'marital_status': 'm',
    'city': 'Wichita',
    'state': '',
    'postal_code': '67226',
    'home_phone': '316-304-4347',
    'other_phone': '',
    'email': '',
    'income': '',
    'tour_date': '7/8',
    'tour_time': '1030',
    'type_of_deposit': '',
    'deposit_amount': '',
    'card_number': '',
    'expiration': '',
    'cvv_code': '',
    'accommodation': '',
    'arrival_date': '',
    'number_of_nights': '',
    'number_of_adults': '',
    'number_of_kids': ''
}

vendor_dict = {
    'location': '5',
    'tour_type': 'm',
    'first_name': 'Bassam',
    'last name': 'Jaradat',
    'spouse_first_name': 'Fatheil',
    'spouse_last_name': 'Abdallah',
    'marital_status': 'm',
    'address': '3540 n inwood st',
    'city': 'Wichita',
    'state': '',
    'postal_code': '67226',
    'phone1': '316-304-4347',
    'phone2': '',
    'email': '',
    'tour_date': '7/8',
    'tour_time': '1030',
    'income': '60',
    'agent_name': 'narancich',
    'campaign': 'bttordm',
    'sales_line': 'ML',
    'resort_code:': 'BR',
    'tm_center': 'BT'
}

switch_site(data_dict)
insert_new_pid()
enter_prospect_info(data_dict)

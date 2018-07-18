import keyboard
import pyautogui

m1 = {}
m2 = {}
m3 = {}
m4 = {}
m5 = {}


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
    m2['insert_tour'] = (m2_title[0] + 473, m2_title[1] + 190)
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
    print(m4_title)


# Record will be Added (Co-prospect menu)
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
    m1 = get_m1_coordinates()
    pyautogui.click(m1['insert'])


def enter_prospect_info(vendor_dict):
    m2 = get_m2_coordinates()
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
    get_m5_coordinates()
    pyautogui.click(m5['get_from_prospect'])
    pyautogui.click(m5['first'])
    keyboard.send(vendor_dict['spouse_first_name'])
    pyautogui.click(m5['ok'])
    pyautogui.click(m2['insert_tour'])
    #  Menu 3 - Adding a Tour Record
    get_m3_coordinates()
    pyautogui.click(m3['campaign'])
    #  Menu 4 - Select a Campaign
    get_m4_coordinates()
    pyautogui.click(m4['clear'])
    pyautogui.click(m4['campaign'])
    keyboard.write(vendor_dict['campaign'])
    pyautogui.click(m4['select'])
    # Menu 3 - Adding a Tour Record
    get_m3_coordinates()
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

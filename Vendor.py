import keyboard
import pyautogui
import time
import pyperclip


m1 = {}
m2 = {}
m3 = {}
m4 = {}


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
    m2_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\adding_a_prospect.png',
                                              region=(514, 245, 889, 566))
    while m2_title is None:
        m2_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles'
                                                  '\\adding_a_prospect.png',
                                                  region=(514, 245, 889, 566))
    m2['title'] = (m2_title[0], m2_title[1])
    # Tabs
    m2['demographics'] = (m2_title[0] + 5, m2_title[1] + 22)
    #  Tours Tab
    m2['first_tour'] = (m2_title[0] + 375, m2_title[1] + 65)
    m2['second_tour'] = (m2_title[0] + 375, m2_title[1] + 78)
    m2['third_tour'] = (m2_title[0] + 375, m2_title[1] + 91)
    m2['insert'] = (m2_title[0] + 475, m2_title[1] + 191)
    m2['yes_change_sites'] = (m2_title[0] + 275, m2_title[1] + 266)
    #  Prospect Tab
    m2['type'] = (m2_title[0] + 285, m2_title[1] + 49)
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
    m2['camp_type'] = (m2_title[0] + 134, m2_title[1] + 400)
    m2['status'] = (m2_title[0] + 134, m2_title[1] + 427)
    #  Demographics Tab
    m2['marital_status'] = (m2_title[0] + 137, m2_title[1] + 84)
    m2['spouse'] = (m2_title[0] + 220, m2_title[1] + 110)
    m2['occupation'] = (m2_title[0] + 137, m2_title[1] + 137)
    m2['income'] = (m2_title[0] + 40, m2_title[1] + 163)
    m2['preferred_language'] = (m2_title[0] + 137, m2_title[1] + 188)
    m2['card_number'] = (m2_title[0] + 192, m2_title[1] + 379)
    m2['expiration'] = (m2_title[0] + 91, m2_title[1] + 432)
    return m2


def get_m3_coordinates():
    global m3
    m3_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\titles\\adding_a_tour.png',
                                              region=(514, 245, 889, 566))
    while m3_title is None:
        m3_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\titles\\adding_a_tour.png',
                                                  region=(514, 245, 889, 566))
    m3['title'] = (m3_title[0], m3_title[1])
    #  Tabs
    m3['tour'] = (m3_title[0] - 45, m3_title[1] + 20)
    m3['user_fields'] = (m3_title[0] + 5, m3_title[1] + 20)
    m3['notes'] = (m3_title[0] + 55, m3_title[1] + 20)
    m3['accommodations'] = (m3_title[0] + 242, m3_title[1] + 24)
    m3['tour_packages'] = (m3_title[0] + 324, m3_title[1] + 24)
    m3['premiums'] = (m3_title[0] + 397, m3_title[1] + 24)
    m3['personnel'] = (m3_title[0] + 225, m3_title[1] + 217)
    #  Tour tab
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
    m3['insert_1'] = (m3_title[0] + 314, m3_title[1] + 439)
    #  Premiums tab
    m3['premium_1'] = (m3_title[0] + 333, m3_title[1] + 64)
    m3['premium_2'] = (m3_title[0] + 333, m3_title[1] + 78)
    m3['premium_3'] = (m3_title[0] + 333, m3_title[1] + 90)
    m3['premium_4'] = (m3_title[0] + 333, m3_title[1] + 102)
    m3['premium_5'] = (m3_title[0] + 333, m3_title[1] + 114)
    m3['premium_6'] = (m3_title[0] + 333, m3_title[1] + 126)
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


def switch_site(site):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\windows_closed.png',
                                           region=(514, 245, 300, 300))
    if image is None:
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


def insert_new_pid():
    m1 = get_m1_coordinates()
    pyautogui.click(m1['insert'])


def enter_prospect_info(data_dict):
    m2 = get_m2_coordinates()
    pyautogui.click(m2['last_name'])
    keyboard.write(data_dict['last name'])
    pyautogui.click(m2['first_name'])
    keyboard.write(data_dict['first_name'])
    pyautogui.click(m2['address'])
    keyboard.write(data_dict['address'])
    pyautogui.click(m2['city'])
    keyboard.write(data_dict['city'])
    pyautogui.click(m2['state'])

    pyautogui.click(m2['postal_code'])
    keyboard.write(data_dict['postal_code'])
    pyautogui.click(m2['country'])
    for i in range(5):
        keyboard.send('u')
    pyautogui.click(m2['phone1'])
    keyboard.write(data_dict['phone1'])
    if data_dict['email'] != '':
        pyautogui.click(m2['email'])
        keyboard.write(data_dict['email'])
    pyautogui.click(m2['demographics'])
    #  Demographics Tab
    pyautogui.click(m2['marital_status'])
    if data_dict['marital_status'] == 'm':
        keyboard.send('m')
    pyautogui.click(m2['spouse'])
    keyboard.write(data_dict['spouse_first_name'] + ' ' + data_dict['spouse_last_name'])
    pyautogui.click(m2['occupation'])
    keyboard.send('e')
    pyautogui.click(m2['income'])
    keyboard.write(data_dict['income'])
    pyautogui.click(m2['preferred_language'])
    keyboard.send('e')
    pyautogui.doubleClick(m2['card_number'])
    keyboard.write(data_dict['card_number'])
    pyautogui.doubleClick(m2['expiration'])
    keyboard.write(data_dict['expiration'])
    keyboard.send('Tab')
    pyautogui.click(m2['insert'])
    #  Menu 3 - Adding a Tour Record
    get_m3_coordinates()
    pyautogui.click(m3['campaign'])
    #  Menu 4 - Select a Campaign
    get_m4_coordinates()
    pyautogui.click(m4['clear'])
    pyautogui.click(m4['campaign'])
    keyboard.write(data_dict['campaign'])
    pyautogui.click(m4['select'])
    # Menu 3 - Adding a Tour Record
    get_m3_coordinates()
    pyautogui.click(m3['tour_type'])
    keyboard.write(data_dict['tour_type'])
    pyautogui.click(m3['tour_status'])
    keyboard.write('b')
    pyautogui.click(m3['tour_date'])
    keyboard.write(data_dict['tour_date'])
    pyautogui.click(m3['wave'])


data_dict = {
    'agent_name': '',
    'location': 'br',
    'campaign': 'OMOWNMM',
    'tour_type': 'm',
    'pid': '',
    'first_name': 'Bassam',
    'last name': 'Jaradat',
    'spouse_first_name': 'Fatheil',
    'spouse_last_name': 'Abdallah',
    'marital_status': 'm',
    'address': '3540 n inwood st',
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

switch_site(data_dict)
insert_new_pid()
enter_prospect_info(data_dict)

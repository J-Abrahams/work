import keyboard
import pyautogui
import time
import pyperclip


def get_m1_coordinates():
    m1 = {}
    m1_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png',
                                              region=(514, 245, 889, 566))
    while m1_title is None:
        m1_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png',
                                                  region=(514, 245, 889, 566))
    m1['search'] = (m1_title[0] + 50, m1_title[1])
    m1['find_now'] = (m1_title[0] + 650, m1_title[1])
    m1['change'] = (m1_title[0] + 400, m1_title[1] + 500)
    m1['insert'] = (m1_title[0] + 300, m1_title[1] + 500)
    return m1


def get_m2_coordinates():
    m2 = {}
    m2_t2 = {}
    m2_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles\\adding_a_prospect.png',
                                              region=(514, 245, 889, 566))
    while m2_title is None:
        m2_title = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\Titles'
                                                  '\\adding_a_prospect.png',
                                                  region=(514, 245, 889, 566))
    m2['first_tour'] = (m2_title[0] + 338, m2_title[1] + 65)
    m2['second_tour'] = (m2_title[0] + 338, m2_title[1] + 78)
    m2['third_tour'] = (m2_title[0] + 338, m2_title[1] + 91)
    m2['yes_change_sites'] = (m2_title[0] + 275, m2_title[1] + 266)
    m2['prospect_id'] = (m2_title[0] + 28, m2_title[1] + 52)
    m2['type'] = (m2_title[0] + 248, m2_title[1] + 52)
    m2['last_name'] = (m2_title[0] + 92, m2_title[1] + 85)
    m2['first_name'] = (m2_title[0] + 234, m2_title[1] + 85)
    m2['salutation'] = (m2_title[0] + 236, m2_title[1] + 110)
    m2['company'] = (m2_title[0] + 234, m2_title[1] + 137)
    m2['address'] = (m2_title[0] + 234, m2_title[1] + 161)
    m2['city'] = (m2_title[0] + 129, m2_title[1] + 223)
    m2['county'] = (m2_title[0] + 234, m2_title[1] + 223)
    m2['state'] = (m2_title[0] + 103, m2_title[1] + 249)
    m2['postal_code'] = (m2_title[0] + 234, m2_title[1] + 249)
    m2['country'] = (m2_title[0] + 137, m2_title[1] + 276)
    m2['phone1'] = (m2_title[0] + 92, m2_title[1] + 302)
    m2['phone2'] = (m2_title[0] + 234, m2_title[1] + 302)
    m2['fax'] = (m2_title[0] + 92, m2_title[1] + 328)
    m2['camp_type'] = (m2_title[0] + 92, m2_title[1] + 399)
    m2['status'] = (m2_title[0] + 92, m2_title[1] + 427)
    m2_t2['marital_status'] = (m2_title[0] + 136, m2_title[1] + 85)
    m2_t2['occupation'] = (m2_title[0] + 234, m2_title[1] + 111)
    m2_t2['income'] = (m2_title[0] + 34, m2_title[1] + 162)
    m2_t2['income'] = (m2_title[0] + 136, m2_title[1] + 188)
    return m2


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


def insert_new_pid(pid_number):
    m1 = get_m1_coordinates()
    pyautogui.click(m1['insert'])


def enter_prospect_info():
    m2 = get_m2_coordinates()
    pyautogui.click(m2['last_name'])
    keyboard.write(data_dict['last name'])
    pyautogui.click(m2['first_name'])
    keyboard.write(data_dict['first_name'])
    pyautogui.click(m2['address'])

    pyautogui.click(m2['city'])

    pyautogui.click(m2['state'])

    pyautogui.click(m2['postal_code'])

    pyautogui.click(m2['country'])

    pyautogui.click(m2['phone1'])
    keyboard.write(data_dict['phone1'])
    if data_dict['email'] != '':
        pyautogui.click(m2['email'])
        keyboard.write(data_dict['email'])
    if data_dict['spouse_first_name'] and data_dict['spouse_last_name'] != '':


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
    'credit_card_#': '',
    'accommodation': 'vvl',
    'arrival_date': '8/5',
    'number_of_nights': '2',
    'number_of_adults': '2',
    'number_of_kids': '0'
}

switch_site(data_dict)
search_pid(data_dict)
enter_prospect_info(data_dict)

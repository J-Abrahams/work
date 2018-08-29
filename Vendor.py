import csv
import time
from tkinter import Tk
import keyboard
import mss
import mss.tools
import pandas as pd
import pyautogui
import clipboard
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10
import confirmations_auto as conf
import datetime
from tabulate import tabulate
import sys
import pickle

import importlib
# importlib.reload(sc)



def switch_site(site_number):
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\windows_closed.png',
                                           region=(514, 245, 300, 300))
    if image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_search_for.png',
                                               region=(514, 245, 889, 566))
        sc.get_m1_coordinates()
        x, y = m1['search']
        pyautogui.click(m1['close'])
    while pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\windows_closed.png',
                                         region=(514, 245, 300, 300)) is None:
        pass
    pyautogui.click(500, 500)
    keyboard.send('F3')
    time.sleep(0.5)
    if site_number == 2:
        keyboard.write('A1')
    elif site_number == 3:
        keyboard.write('A2')
    elif site_number == 4:
        keyboard.write('A3')
    elif site_number == 5:
        keyboard.write('t')
    elif site_number == 8:
        keyboard.write('c')
    elif site_number == 9:
        keyboard.write('Welk Resort N')
    elif site_number == 11:
        keyboard.write('Welk Resort Bre')
    # time.sleep(1)
    keyboard.send('enter')
    pyautogui.click(10, 25)
    pyautogui.click(10, 45)


def take_screenshot(y, x, width, height, save_file=False):
    with mss.mss() as sct:
        monitor = {'top': y, 'left': x, 'width': width, 'height': height}
        im = sct.grab(monitor)
        screenshot = str(mss.tools.to_png(im.rgb, im.size))
        if save_file:
            now = datetime.datetime.now()
            output = now.strftime("%m-%d-%H-%M-%S.png".format(**monitor))
            number = 1
            output = output + str(number)
            mss.tools.to_png(im.rgb, im.size, output=output)
            number += 1
        return screenshot


def read_pickle_file(file_name):
    with open('text_files\\' + file_name, 'rb') as file:
        return pickle.load(file)


def pause(message):
    print(message)
    while pyautogui.position() != (0, 1079):
        pass


def insert_new_pid():
    sc.get_m1_coordinates()
    pyautogui.click(m1['insert'])


def convert_excel_to_dataframe():
    xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\Downloads\\Vendor.xlsx")
    df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
    print(tabulate(df, headers='keys', tablefmt='psql'))
    return df


def enter_m2_info():
    df = convert_excel_to_dataframe()
    switch_site(df.loc[0, 'Site'])
    insert_new_pid()
    sc.get_m2_coordinates(True)
    pyautogui.click(m2['last_name'])
    keyboard.write(df.loc[0, 'Last_Name'])
    pyautogui.click(m2['first_name'])
    keyboard.write(df.loc[0, 'First_Name'])
    pyautogui.click(m2['address'])
    keyboard.write(df.loc[0, 'Address'])
    pyautogui.click(m2['city'])
    keyboard.write(df.loc[0, 'City'])
    pyautogui.click(m2['state'])

    pyautogui.click(m2['postal_code'])
    keyboard.write(str(df.loc[0, 'Zip']))
    pyautogui.click(m2['country'])
    for i in range(5):
        keyboard.send('u')
    pyautogui.click(m2['phone1'])
    keyboard.write(df.loc[0, 'Phone'])
    """if df.loc[0, 'Phone_2'] != '':
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
                                                       region=(514, 245, 889, 566)))"""


enter_m2_info()
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

#switch_site('5')
#insert_new_pid()
#enter_m2_info(data_dict)

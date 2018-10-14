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
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10, m11
import logging
import datetime
from tabulate import tabulate
import sys
import pickle
import openpyxl
import re
import core_functions as cf
import sqlite3


def convert_excel_to_csv():
    xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\Downloads\\Book11.xlsx")
    df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
    df.to_csv('file.csv')


def get_date(month, day, year):
    f = open('text_files\\dates.p', 'rb')
    date_dictionary = pickle.load(f)
    f.close()
    try:
        month = date_dictionary[month]
        day = date_dictionary[day]
        year = date_dictionary[year]
        if month not in ['Nothing', 'Error']:
            tour_date = ('{}/{}/{}'.format(month, day, year))
            return tour_date
            # datetime.datetime.strptime(tour_date, "%m/%d/%Y")
        elif month == 'Error':
            return 'Error'
        elif month == 'Nothing':
            return 'Nothing'
    except KeyError:
        return 'Nothing'


def create_data_frame():
    """
    Takes screenshots of the tours. Turns the screenshots into a list of dictionaries 'd'. Turns 'd' into a dataframe.
    d is a list of dictionaries such as [{''Date': '7/06/18', 'Tour_Type': 'Audition', 'Tour_Status': 'Showed'},
    {'Date': '7/06/18', 'Tour_Type': 'minivac', 'Tour_Status': 'Showed'}]
    :return:
    """
    sc.get_m2_coordinates()
    d = []
    pretty_d = []
    x, y = m2['title']
    for i in range(8):
        month = cf.take_screenshot(x + 327, y + 63, 13, 10)
        day = cf.take_screenshot(x + 342, y + 63, 15, 10)
        year = cf.take_screenshot(x + 358, y + 63, 27, 10)
        tour_date = get_date(month, day, year)
        tour_type = cf.take_screenshot(x + 402, y + 63, 14, 10)
        tour_status = cf.take_screenshot(x + 484, y + 63, 14, 10)
        try:
            tour_type = sc.m2_tour_types[tour_type]
        except KeyError:
            print('Unrecognized tour type')
            print(tour_type)
            tour_type = None
        try:
            tour_status_dict = cf.read_pickle_file('m2_tour_status.p')
            tour_status = tour_status_dict[tour_status]
        except KeyError:
            print('Unrecognized tour status')
            print(tour_status)
            tour_status = None
        y += 13
        if tour_date != 'Nothing':
            try:
                # Where the screenshots get turned into dictionaries.
                pretty_d.append({'Date': tour_date, 'Tour_Type': tour_type, 'Tour_Status': tour_status})
                tour_date = pd.to_datetime(tour_date)
                d.append({'Date': tour_date, 'Tour_Type': tour_type, 'Tour_Status': tour_status})
            except NameError:
                pass
    df = pd.DataFrame(d)  # Turn d into a dataframe
    df = df[['Date', 'Tour_Type', 'Tour_Status']]  # Reorders the columns in the dataframe.
    pretty_df = pd.DataFrame(pretty_d)
    pretty_df = pretty_df[['Date', 'Tour_Type', 'Tour_Status']]
    #  print(tabulate(pretty_df, headers='keys', tablefmt='psql'))
    return df


def select_tour(df, attempt_number):
    x, y = m2['title']
    # Returns the top tour that is Showed, not an Audition, and at most a week before the date we entered.
    # tour_number is the index of the correct tour. Ex: 1 if the second tour is the correct one.
    tour_number = df[(df.Tour_Type != 'Audition')].index[attempt_number - 1]
    pyautogui.doubleClick(x + 469, y + 67 + 13 * tour_number)
    # Checks if "You need to change sites" message comes up
    time.sleep(1)
    pyautogui.click(m2['yes_change_sites'])
    pyautogui.doubleClick(x + 469, y + 67 + 13 * tour_number)
    screen_shot = None
    while screen_shot == sc.no_popup or screen_shot is None:
        with mss.mss() as sct:
            monitor = {'top': 507, 'left': 941, 'width': 23, 'height': 13}
            im = sct.grab(monitor)
            screen_shot = str(mss.tools.to_png(im.rgb, im.size))
    if screen_shot != sc.no_incorrect_site:
        pyautogui.click(m2['yes_change_sites'])


convert_excel_to_csv()
with open('file.csv') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        pid = row['PID'].replace('.0', '')
        cf.search_pid(pid)
        df = create_data_frame()
        select_tour(df, 1)
        sc.get_m3_coordinates()
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\a2.png',
                                               region=(514, 245, 889, 566))
        while image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\a2.png',
                                                   region=(514, 245, 889, 566))
        pyautogui.click(image)
        keyboard.send('g')
        pyautogui.click(m3['wave'])
        print(row['Wave'])
        if row['Wave'] == '0':
            pass
        elif row['Wave'] == '900':
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_900.png',
                                                   region=(514, 245, 889, 566))
            while image is None:
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_900.png',
                                                       region=(514, 245, 889, 566))
            pyautogui.click(image)
        elif row['Wave'] == '930':
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_930.png',
                                                   region=(514, 245, 889, 566))
            while image is None:
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_930.png',
                                                       region=(514, 245, 889, 566))
            pyautogui.click(image)
        elif row['Wave'] == '1200':
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1200.png',
                                                   region=(514, 245, 889, 566))
            while image is None:
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1200.png',
                                                       region=(514, 245, 889, 566))
            pyautogui.click(image)
        elif row['Wave'] == '1230':
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1230.png',
                                                   region=(514, 245, 889, 566))
            while image is None:
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1230.png',
                                                       region=(514, 245, 889, 566))
            pyautogui.click(image)
        elif row['Wave'] == '1430':
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1430.png',
                                                   region=(514, 245, 889, 750))
            while image is None:
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1430.png',
                                                       region=(514, 245, 889, 750))
            pyautogui.click(image)
        elif row['Wave'] == '1500':
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1500.png',
                                                   region=(514, 245, 889, 566))
            while image is None:
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_1500.png',
                                                       region=(514, 245, 889, 566))
            pyautogui.click(image)
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
        while image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                                   region=(514, 245, 889, 566))
        x, y = image
        pyautogui.click(x + 265, y + 475)
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                               region=(514, 245, 889, 566))
        while image is None:
            image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                                   region=(514, 245, 889, 566))
        x, y = image
        pyautogui.click(x - 20, y + 425)
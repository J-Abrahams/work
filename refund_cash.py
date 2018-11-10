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
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10, m11, m12, m13, m14
import logging
import datetime
from tabulate import tabulate
import sys
import pickle
import openpyxl
import re
import core_functions as cf
import sqlite3
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from collections import Counter

conn = sqlite3.connect('sqlite.sqlite')


def read_pickle_file(file_name):
    with open('text_files\\' + file_name, 'rb') as file:
        return pickle.load(file)


def sqlite_cursor():
    c = sqlite3.connect('sqlite.sqlite').cursor()
    return c


def sqlite_get_item(sqlite_statement, *parameters):
    global conn
    c = conn.cursor()
    c.execute(sqlite_statement, *parameters)
    return c.fetchone()[0]


def sqlite_select(screenshot, table):
    conn = sqlite3.connect('sqlite.sqlite')
    c = conn.cursor()
    c.execute("SELECT {} FROM {} WHERE screenshot=?".format('name', table), [screenshot])
    try:
        name = c.fetchone()[0]
    except TypeError:
        if table == 'numbers':
            name = ''
        elif table == 'premiums':
            name = 'Old Premium'
    return name


def select_tour():
    df, pretty_df = cf.create_accommodations_dataframe()
    # print(tabulate(pretty_df, headers='keys', tablefmt='psql'))
    # Returns the top tour that is Showed, not an Audition, and at most a week before the date we entered.
    # tour_number is the index of the correct tour. Ex: 1 if the second tour is the correct one.
    try:
        tour_number = df[(df.Tour_Status == 'No_Tour') & (df.Tour_Type != 'Audition')].index[0]
        x, y = m2['title']
        pyautogui.doubleClick(x + 469, y + 67 + 13 * tour_number)

        # Checks if "You need to change sites" message comes up
        # Messed up here and had to make it confusing
        screen_shot = None
        while screen_shot == sc.no_popup or screen_shot is None:
            with mss.mss() as sct:
                monitor = {'top': 507, 'left': 941, 'width': 23, 'height': 13}
                im = sct.grab(monitor)
                screen_shot = str(mss.tools.to_png(im.rgb, im.size))
        if screen_shot != sc.no_incorrect_site:
            pyautogui.click(m2['yes_change_sites'])
            return
    except IndexError:
        pass
        print('Couldn\'t find correct tour')
        tour_number = df[(df.Tour_Type != 'Audition') & (df.Tour_Status != 'Error')].index[0]
    try:
        tour_number = df[(df.Tour_Type != 'Audition')].index[0]
        x, y = m2['title']
        pyautogui.doubleClick(x + 469, y + 67 + 13 * tour_number)

        # Checks if "You need to change sites" message comes up
        # Messed up here and had to make it confusing
        screen_shot = None
        while screen_shot == sc.no_popup or screen_shot is None:
            with mss.mss() as sct:
                monitor = {'top': 507, 'left': 941, 'width': 23, 'height': 13}
                im = sct.grab(monitor)
                screen_shot = str(mss.tools.to_png(im.rgb, im.size))
        if screen_shot != sc.no_incorrect_site:
            pyautogui.click(m2['yes_change_sites'])
            return
    except IndexError:
            pass
    try:
        tour_number = df[(df.Tour_Status == 'No_Tour') & (df.Tour_Type != 'Audition')].index[0]
        x, y = m2['title']
        pyautogui.doubleClick(x + 469, y + 67 + 13 * tour_number)

        # Checks if "You need to change sites" message comes up
        # Messed up here and had to make it confusing
        screen_shot = None
        while screen_shot == sc.no_popup or screen_shot is None:
            with mss.mss() as sct:
                monitor = {'top': 507, 'left': 941, 'width': 23, 'height': 13}
                im = sct.grab(monitor)
                screen_shot = str(mss.tools.to_png(im.rgb, im.size))
        if screen_shot != sc.no_incorrect_site:
            pyautogui.click(m2['yes_change_sites'])
            return
    except IndexError:
            pass
    x, y = m2['title']
    pyautogui.doubleClick(x + 469, y + 67 + 13 * tour_number)

    # Checks if "You need to change sites" message comes up
    # Messed up here and had to make it confusing
    screen_shot = None
    while screen_shot == sc.no_popup or screen_shot is None:
        with mss.mss() as sct:
            monitor = {'top': 507, 'left': 941, 'width': 23, 'height': 13}
            im = sct.grab(monitor)
            screen_shot = str(mss.tools.to_png(im.rgb, im.size))
    if screen_shot != sc.no_incorrect_site:
        pyautogui.click(m2['yes_change_sites'])


def select_premium(balance, index):
    premium_list = []
    sc.get_m3_coordinates()
    x, y = m3['title']
    pyautogui.click(m3['premiums'])
    pyautogui.click(m3['notes'])
    # action = input('What to do?')
    # if action not in ['note', '50', '40']:
    #     return 'good'
    # elif action == 'note':
    #     return 'note'
    # else:
    #     balance = action
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\issued.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\issued.png',
                                               region=(514, 245, 889, 566))
    while True:
        screenshot_2 = cf.take_screenshot_change_color(x + 433, y + 60, 10, 8)
        premium = sqlite_select(cf.take_screenshot_change_color(x + 342, y + 60, 80, 11), 'premiums')
        if premium == 'nothing':
            break
        elif screenshot_2 == sqlite_get_item("SELECT screenshot FROM premiums WHERE name=?", ['Did Not Issue']):
            premium = f'{premium} - Canceled'
            premium_list.append(premium)
            y += 13
            continue
        else:
            premium_list.append(premium)
            y += 13
    number_of_dep_premiums_50 = premium_list.count('DEP $50 CC') + premium_list.count('$50 CC Dep')
    number_of_dep_premiums_40 = premium_list.count('DEP $40 CC') + premium_list.count('$40 CC Dep')
    total_number_of_dep_premiums = number_of_dep_premiums_50 + number_of_dep_premiums_40
    if balance == '0' and total_number_of_dep_premiums == 0:
        return 'no note'
    elif balance != '0' and total_number_of_dep_premiums == 0:
        sheet.update_cell(index, 4, 'No DEPs')
        return 'no note'
    elif total_number_of_dep_premiums > 1:
        sheet.update_cell(index, 4, 'Multiple DEPs')
        return 'no note'
    elif total_number_of_dep_premiums == '0' and f"DEP ${balance} Cash" in premium_list:
        sheet.update_cell(index, 4, 'Already done')
        return 'no note'
    else:
        x, y = m3['title']
        for premium in premium_list:
            if premium in ['DEP $50 CC', '$50 CC Dep', 'DEP $40 CC', '$40 CC Dep'] and balance in premium:
                pyautogui.click(x + 342, y + 60)
                pyautogui.click(sc.m3['change_premium'])
                sc.get_m10_coordinates()
                x, y = sc.m10['title']
                pyautogui.click(x + 198, y + 423)
                pyautogui.dragTo(x - 100, y + 300, button='left')
                keyboard.send('backspace')
                keyboard.write('Please refund by cash the charge is over 120 days and cannot be refunded back to CC.')
                pyautogui.click(x + 160, y + 80)
                image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\titles'
                                                       '\\premium_search.png',
                                                       region=(514, 245, 889, 566))
                while image is None:
                    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\titles'
                                                           '\\premium_search.png',
                                                           region=(514, 245, 889, 566))
                x, y = image
                pyautogui.click(x + 189, y + 35)
                keyboard.write(f"Dep ${balance} cash")
                pyautogui.click(x + 267, y + 35)
                pyautogui.click(x + 129, y + 476)
                sc.get_m10_coordinates()
                x, y = sc.m10['title']
                pyautogui.click(x + 80, y + 503)
                return
            else:
                y += 13


def add_note():
    sc.get_m3_coordinates()
    x, y = m3['title']
    pyautogui.click(x, y + 435)
    sc.get_m14_coordinates()
    sites_dictionary = read_pickle_file('sites.p')
    site = cf.take_screenshot(1517, 1036, 146, 17)
    if sites_dictionary[site] == 'Northstar':
        keyboard.write(' Blank')
    else:
        keyboard.write('Blank')
    pyautogui.click(m14['note'])
    keyboard.write('Please refund by cash the charge is over 120 days and cannot be refunded back to CC.')
    pyautogui.click(m14['ok'])


def show_progress(pid, progress, number_of_pids):
    percentage = round(progress * 100 / number_of_pids, 2)
    print('{} / {} - {} - {}'.format(str(progress), str(number_of_pids), str(percentage) + '%', str(pid)))


cf.pause('Minimize Pycharm')
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('Phone-6ad41718c799.json', scope)
client = gspread.authorize(creds)
sheet = client.open("Cash Refund").sheet1
dictionaries = sheet.get_all_records()
number_of_rows = len(sheet.get_all_records())
progress = 1
index = 1
for row in dictionaries:
    status = []
    index += 1
    pid, balance, completed = str(row['pid']), str(row['balance']), row['completed']
    if completed in ['x', 'X']:
        # if balance not in ['50', '40'] or completed in ['x', 'X']:
        progress += 1
        continue
    show_progress(pid, progress, number_of_rows)
    cf.search_pid(pid)
    select_tour()
    found_premium = select_premium(balance, index)
    if found_premium != 'no note':
        add_note()
    sc.get_m3_coordinates()
    pyautogui.click(m3['ok'])
    sc.get_m2_coordinates()
    pyautogui.click(m2['ok'])
    sheet.update_cell(index, 3, 'x')
    progress += 1

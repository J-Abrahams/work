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
import datetime
from tabulate import tabulate
import sys
import pickle
import openpyxl
import re
import core_functions as cf
import pytest
import logging
import sqlite3
from oauth2client.service_account import ServiceAccountCredentials
import gspread

sol_numbers = {'Jennifer Gordon': 'SOL2956', 'Katherine England': 'SOL23521', 'Katherine Albini': 'SOL23521',
               'Katherine England/Abini': 'SOL23521', 'Justin Locke': 'SOL4967', 'Brian Bennett': 'SOL3055',
               'Carter Roedell': 'SOL23345', 'Fernanda Hernandez': 'SOL26788', 'Fern Hernandez': 'SOL26788',
               'Alton Major': 'SOL4809', 'Thuy Pham': 'SOL25688', 'Julianne Martinez': 'SOL22766',
               'Quenton Stroud': 'SOL27228', 'Sadie Oliver': 'SOL26834', 'Valeria Rebollar': 'SOL24218',
               'Sergio Espinoza': 'SOL23542', 'Olivia Larimer': 'SOL5463', 'Grayson Corbin': 'SOL1604',
               'Deonte Keller': 'SOL27498', 'Rayven Alexander': 'SOL24125', 'Deeandra Castillo': 'SOL5495',
               'Kenan Williams': 'SOL27567', 'Jenniffer Abbott': 'SOL5456', 'Met Austin Simon': 'SOL27647',
               'Dana Durant': 'SOL27561', 'Seo Ra Yoo': 'SOL27551'}


conn = sqlite3.connect('sqlite.sqlite')


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


def sqlite_get_item(sqlite_statement, *parameters):
    global conn
    c = conn.cursor()
    c.execute(sqlite_statement, *parameters)
    return c.fetchone()[0]


def open_sheet(sheet_name='Confirmation Sheet'):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('Phone-6ad41718c799.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).sheet1
    return sheet


def count_rows(sheet):
    dictionaries = sheet.get_all_records()
    number_of_rows = len(sheet.get_all_records())
    return number_of_rows, dictionaries


def select_tour(tours_list, row):
    x, y = m2['title']
    current_date = cf.get_current_date()
    correct_tour = None
    for tour in tours_list:
        if row.c in ['x', 'X'] and tour.type != 'Audition' \
                and tour.status in ['Showed', 'Confirmed', 'No_Show', 'On_Tour'] \
                and ((tour.date - current_date) >= datetime.timedelta(days=-1)
                     or (tour.date - current_date) <= datetime.timedelta(days=14)):
            pyautogui.doubleClick(x + 469, y + 67 + 13 * tour.index)
            correct_tour = tour
            break
        elif row.r in ['x', 'X'] and (tour.type == 'Open_Reservation'
                                      and tour.date == datetime.datetime.strptime('1/1/1900', "%m/%d/%Y")) \
                or (tour.status == 'Rescheduled' and (tour.date - current_date) >= datetime.timedelta(days=0)) \
                and tour.type != 'Audition':
            pyautogui.doubleClick(x + 469, y + 67 + 13 * tour.index)
            correct_tour = tour
            break
        elif row.x in ['x', 'X'] and tour.status == 'Canceled' and tour.type != 'Audition':
            pyautogui.doubleClick(x + 469, y + 67 + 13 * tour.index)
            correct_tour = tour
            break
        elif row.u in ['x', 'X'] and tour.type == 'Minivac' and ((tour.date - current_date) >= datetime.timedelta(days=-1)):
            pyautogui.doubleClick(x + 469, y + 67 + 13 * tour.index)
            correct_tour = tour
            break
        elif row.t in ['x', 'X'] and tour.type == 'Day_Drive' \
                and ((tour.date - current_date) >= datetime.timedelta(days=-1)):
            pyautogui.doubleClick(x + 469, y + 67 + 13 * tour.index)
            correct_tour = tour
            break
    if correct_tour is None:
        for tour in tours_list:
            if tour.type != 'Audition' and tour.status != 'Error':
                print('Couldn\'t find correct tour')
                pyautogui.doubleClick(x + 469, y + 67 + 13 * tour.index)
                correct_tour = tour
                break
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
    return correct_tour


def count_accommodations():
    number_of_accommodations = 0
    number_of_canceled_accommodations = 0
    x, y = m3['title']
    while True:
        screenshot = cf.take_screenshot_change_color(x + 587, y + 64, 6, 9)
        if screenshot == 'nothing':
            return number_of_accommodations, number_of_canceled_accommodations
        else:
            screenshot_2 = cf.take_screenshot_change_color(x + 272, y + 64, 47, 8)
            if screenshot_2 == 'nothing':
                number_of_canceled_accommodations += 1
                y += 13
            else:
                number_of_accommodations += 1
                y += 13


def count_deposits():
    number_of_deposits = 0
    pyautogui.click(m3['tour_packages'])
    x, y = m3['title']
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\balance.png',
                                           region=(700, 245, 850, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\balance.png',
                                               region=(700, 245, 850, 566))
    while True:
        # Counts number of deposits.
        # Breaks 'while' loop once a returned screenshot is blank
        deposit_screenshot = cf.take_screenshot(x + 255, y + 69, 6, 9)
        if deposit_screenshot == sc.no_deposits:
            return number_of_deposits
        else:
            number_of_deposits += 1
            y += 13


def count_premiums():
    number_of_premiums = 0
    sc.get_m3_coordinates()
    x, y = m3['title']
    pyautogui.click(m3['premiums'])
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\issued.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\issued.png',
                                               region=(514, 245, 889, 566))
    while True:
        premium = cf.take_screenshot_change_color(x + 342, y + 60, 80, 11)
        if premium == 'nothing':
            break
        else:
            number_of_premiums += 1
            y += 13
    return number_of_premiums


class ConfirmationSheet:

    def __init__(self, sheet_name):
        self.sheet = self.open_sheet(sheet_name)

    def open_sheet(self, sheet_name):
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('Phone-6ad41718c799.json', scope)
        client = gspread.authorize(creds)
        sheet = client.open(sheet_name).sheet1
        return sheet

    def read_rows(self):
        list_of_rows = self.sheet.get_all_records()
        return list_of_rows

    def count_rows(self):
        number_of_rows = len(self.sheet.get_all_records())
        return number_of_rows


class Row:

    def __init__(self, pid, c, r, x, u, t, completed):
        self.pid = str(pid)
        self.c = c
        self.r = r
        self.x = x
        self.u = u
        self.t = t
        self.completed = completed

    def search_pid(self):
        sc.get_m1_coordinates()
        pyautogui.doubleClick(m1['search'])
        keyboard.write(self.pid)
        pyautogui.click(m1['find_now'])
        pyautogui.click(m1['change'])

    def double_check_pid(self):
        sc.get_m2_coordinates()
        x, y = m2['title']
        pid_screenshot = ''
        conn = sqlite3.connect('sqlite.sqlite')
        c = conn.cursor()
        for i in range(7):
            screenshot = cf.take_screenshot_change_color(x + 74 + 6 * i, y + 48, 6, 9)
            c.execute("SELECT name FROM numbers WHERE screenshot=?", [screenshot])
            number = str(c.fetchone()[0])
            if number != 'nothing':
                pid_screenshot += number
            else:
                continue
        if pid_screenshot != self.pid:
            cf.pause('Is the pid correct?')
            return


class M2Tour:

    def __init__(self, row, tour_number):
        self.row = row
        self.index = tour_number
        self.date = self.tour_date()
        self.type = self.tour_type()
        self.status = self.tour_status()
        self.site = self.site()

    def tour_date(self):
        x, y = m2['title']
        month = cf.take_screenshot(x + 327, y + 63 + self.index * 13, 13, 10)
        day = cf.take_screenshot(x + 342, y + 63 + self.index * 13, 15, 10)
        year = cf.take_screenshot(x + 358, y + 63 + self.index * 13, 27, 10)
        date_str = cf.turn_screenshots_into_date(month, day, year)
        if date_str != 'Nothing':
            return pd.to_datetime(date_str)
        else:
            return pd.to_datetime('1/1/2000')

    def tour_type(self):
        x, y = m2['title']
        tour_type = cf.take_screenshot(x + 402, y + 63 + self.index * 13, 14, 10)
        try:
            tour_type = sc.m2_tour_types[tour_type]
        except KeyError:
            print('Unrecognized tour type')
            print(tour_type)
            tour_type = None
        return tour_type

    def tour_status(self):
        x, y = m2['title']
        tour_status = cf.take_screenshot(x + 484, y + 63 + self.index * 13, 14, 10)
        try:
            tour_status_dict = cf.read_pickle_file('m2_tour_status.p')
            tour_status = tour_status_dict[tour_status]
        except KeyError:
            print('Unrecognized tour status')
            print(tour_status)
            tour_status = None
        return tour_status

    def site(self):
        sc.get_m2_coordinates()
        x, y = m2['title']
        conn = sqlite3.connect('sqlite.sqlite')
        c = conn.cursor()
        if pyautogui.pixelMatchesColor(1318, 358, (214, 211, 206)):
            screenshot = cf.take_screenshot_change_color(x + 671, y + 64 + self.index * 13, 6, 9)
        else:
            screenshot = cf.take_screenshot_change_color(x + 687, y + 64 + self.index * 13, 6, 9)
        if screenshot != 'nothing':
            c.execute("SELECT name FROM numbers WHERE screenshot=?", [screenshot])
            tour_site = str(c.fetchone()[0])
            return tour_site


class Tour:

    def __init__(self, chosen_tour, row):
        self.campaign = None
        self.type = chosen_tour.type
        self.status = chosen_tour.status
        self.date = chosen_tour.date
        self.location = None
        self.wave = self.wave()
        self.site = chosen_tour.site

    def wave(self):
        sc.get_m3_coordinates()
        x, y = m3['title']
        conn = sqlite3.connect('sqlite.sqlite')
        c = conn.cursor()
        wave = ''
        for i in range(4):
            screenshot = cf.take_screenshot_change_color(x + 37 + 6 * i, y + 247, 6, 9)
            if screenshot != 'nothing':
                c.execute("SELECT name FROM numbers WHERE screenshot=?", [screenshot])
                number = str(c.fetchone()[0])
                wave += number
            else:
                continue
        return wave


class Deposit:

    def __init__(self, index):
        self.index = index
        self.description = self.description()
        self.type = self.deposit_type()
        self.amount = self.amount()

    def description(self):
        x, y = m3['title']
        screenshot = cf.take_screenshot_change_color(x + 266, y + 68 + 13 * self.index, 154, 10)
        conn = sqlite3.connect('sqlite.sqlite')
        c = conn.cursor()
        try:
            c.execute("SELECT name FROM deposits WHERE screenshot=?", [screenshot])
            description = c.fetchone()[0]
            return description
        except TypeError:
            clipboard.copy('bad')
            pyautogui.click(x + 266, y + 68 + 13 * self.index)
            pyautogui.click(m3['change_deposit'])
            sc.get_m6_coordinates()
            pyautogui.click(m6['description'])
            keyboard.send('ctrl + z')
            keyboard.send('ctrl + c')
            r = Tk()
            description = r.selection_get(selection="CLIPBOARD")
            while description == 'bad':
                pyautogui.click(m6['description'])
                keyboard.send('ctrl + z')
                keyboard.send('ctrl + c')
                description = r.selection_get(selection="CLIPBOARD")
            pyautogui.click(m6['ok'])
            return description

    def deposit_type(self):
        if 'refunded' in self.description.lower():
            deposit_type = 'refunded'
        elif 'minivac' in self.description.lower() or 'apply' in self.description.lower():
            deposit_type = 'non_refundable'
        elif 'ref' in self.description.lower():
            deposit_type = 'refundable'
        else:
            deposit_type = 'non_refundable'
        return deposit_type

    def amount(self):
        x, y = m3['title']
        hundreds = str(sqlite_select(cf.take_screenshot_change_color(x + 467, y + 69 + self.index * 13, 6, 9), 'numbers'))
        tens = str(sqlite_select(cf.take_screenshot_change_color(x + 473, y + 69 + self.index * 13, 6, 9), 'numbers'))
        ones = str(sqlite_select(cf.take_screenshot_change_color(x + 479, y + 69 + self.index * 13, 6, 9), 'numbers'))
        if hundreds == 'nothing':
            hundreds = ''
        if tens == 'nothing':
            tens = ''
        amount = hundreds + tens + ones
        return amount


class Premium:

    def __init__(self, index):
        self.index = index
        self.name = self.name()
        self.canceled = self.canceled()
        self.refundable = self.refundable()

    def name(self):
        x, y = m3['title']
        premium = sqlite_select(cf.take_screenshot_change_color(x + 342, y + 60 + self.index * 13, 80, 11), 'premiums')
        return premium

    def canceled(self):
        global conn
        c = conn.cursor()
        x, y = m3['title']
        screenshot = cf.take_screenshot_change_color(x + 433, y + 60 + self.index * 13, 10, 8)
        # c.execute("SELECT screenshot FROM premiums WHERE name=?", ['Did Not Issue'])
        # did_not_issue = c.fetchone[0]
        if screenshot == 'nothing':
            return 'no'
        else:
            return 'yes'

    def refundable(self):
        if 'DEP' in self.name:
            return 'yes'
        else:
            return 'no'

if __name__ == "__main__":
    confirmation_sheet = ConfirmationSheet("Confirmation Sheet")

    # Count number of rows
    list_of_rows = confirmation_sheet.read_rows()
    number_of_rows = confirmation_sheet.count_rows()

    for row in list_of_rows:

        # Changes the sol number if the column 'Sol' is not empty.
        if row['Sol'] != '':
            sol = sol_numbers[row['Sol']]

        # Create a row object.
        row = Row(row['PID'], row['conf'], row['rxl'], row['cxl'], row['ug'], row['tav'], row['Completed'])

        # Skips row if the completed column is checked off.
        if row.completed in ['x', 'X']:
            continue

        # Enters, searches, and selects the PID.
        row.search_pid()

        # Checks that the correct PID was selected.
        row.double_check_pid()
        # For each tour that is listed for the PID on M2, creates an instance of M2Tour and adds it to the list tours.
        tours = []
        for i in range(8):
            tour = M2Tour(row, i)
            if tour.date == pd.to_datetime('1/1/2000') and tour.type == '' and tour.status == '':
                break
            else:
                tours.append(tour)
                print(tour.date, tour.type, tour.status, tour.site)
        # Selects correct tour
        correct_tour = select_tour(tours, row)

        # Creates a tour object with all the face info.
        sc.get_m3_coordinates()
        tour = Tour(correct_tour, row)

        # Count number of Accommodations
        number_of_accommodations, number_of_canceled_accommodations = count_accommodations()

        # Count number of Deposits
        number_of_deposits = count_deposits()
        deposits = []
        deposits_df = []
        for i in range(number_of_deposits):
            deposit = Deposit(i)
            deposits.append(deposit)
            deposits_df.append({'description': deposit.type, 'amount': deposit.amount})
        deposits_df = pd.DataFrame(deposits_df)
        print(deposits_df)
        print(tour.status, tour.type, tour.wave)
        number_of_refundable_deposits = 0
        for deposit in deposits:
            if deposit.type == 'refundable':
                number_of_refundable_deposits += 1
            print('1', deposit.amount, deposit.type)

        # Count number of Premiums
        number_of_premiums = count_premiums()
        premiums = []
        premiums_df = []
        for i in range(number_of_premiums):
            premium = Premium(i)
            premiums.append(premium)
            premiums_df.append({'premium': premium.name, 'refundable': premium.refundable, 'canceled': premium.canceled})
        premiums_df = pd.DataFrame(premiums_df)
        print(premiums_df)
        for premium in premiums:
            print(premium.name, premium.canceled, premium.refundable)

        if (tour.status in ['canceled', 'no_tour'] and number_of_accommodations == 0) or \
                (tour.type == 'Minivac' and number_of_accommodations > 0) or \
                (tour.type == 'Day_drive' and number_of_accommodations == 0):
            cf.print_colored_text(f'{tour.type} - {str(number_of_accommodations)}', 'green')
        else:
            cf.print_colored_text(f'{tour.type} - {str(number_of_accommodations)}', 'red')
        if len(premiums) != len(set(premiums)):
            cf.print_colored_text(f'{len(premiums)} Premium(s) - DUPLICATES', 'red')
        else:
            cf.print_colored_text(f'{len(premiums)} Premium(s) - No Duplicates', 'green')
        if len(deposits_df) != 0 and len(premiums_df) != 0:
            if len(deposits_df[deposits_df['description'] == 'refundable']) == len(premiums_df[premiums_df['refundable'] == 'yes']):
                print('good')
        # for premium in premiums:
        #     if premium.refundable is True and premium.canceled is False:
        #         pass
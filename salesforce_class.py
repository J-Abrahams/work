from simple_salesforce import Salesforce
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
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10, m11, m12, m13
import logging
import datetime
from tabulate import tabulate
import sys
import pickle
import re
import core_functions as cf
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


def select_tour(tours_list, row):
    x, y = m2['title']
    current_date = cf.get_current_date()
    correct_tour = None
    for tour in tours_list:
        if tour.status == 'Showed' and tour.type != 'Audition' and ((row.date - tour.date) <= datetime.timedelta(days=7)):
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


class ReflectionsSheet:

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

    def __init__(self, pid, price, date, cash, completed):
        self.pid = str(pid)
        self.price = price
        self.date = datetime.datetime.strptime(date + '/18', "%m/%d/%y")
        self.cash = cash
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


class Opportunity:

    def __init__(self, opportunity, index):
        self.index = index
        self.tid = opportunity['TourId__c']
        self.date = self.date()
        self.type = opportunity['TourType__c']
        self.status = opportunity['Disposition__c']
        self.site = opportunity['SiteID2__c']

    def date(self):
        date = pd.to_datetime(opportunity['Tour_Date__c'])
        return date.date()


if __name__ == "__main__":
    sf = Salesforce(username='jared.abrahams@welkgroup.com', password='Guybrush2',
                    security_token='LNM0x9TuTtxAsccksDSAWaNi')
    progress = 0
    reflections_sheet = ReflectionsSheet("Reflections")

    # Count number of rows
    list_of_rows = reflections_sheet.read_rows()
    number_of_rows = reflections_sheet.count_rows()

    for row in list_of_rows:
        # Create a row object.
        row = Row(str(row['pid']), row['price'], row['date'], row['cash'], row['completed'])

        # Skips row if the completed column is checked off.
        if row.completed in ['x', 'X']:
            progress += 1
            continue

        # Enters, searches, and selects the PID.
        row.search_pid()

        # Checks that the correct PID was selected.
        row.double_check_pid()

        # Queries Salesforce
        opportunities = sf.query(f"SELECT TourId__c ,Tour_Date__c, TourType__c, StageName, Disposition__c, "
                                 f"SiteID2__c, Premium_1__c "
                                 f"FROM Opportunity WHERE ProspectID__c = '{row.pid}' ORDER BY TourID__c DESC")
        tours_list = []
        index = 0
        for opportunity in opportunities["records"]:
            print(opportunity['Premium_1__c'])
            tour = Opportunity(opportunity, index)
            index += 1
            print(tour.date)
            if len(tour.tid) < 7:
                continue
            else:
                tours_list.append(tour)
        correct_tour = select_tour(tours_list, row)
        print(correct_tour.index)

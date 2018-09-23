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


sol_numbers = {'Jennifer Gordon': 'SOL2956', 'Katherine England': 'SOL23521', 'Katherine Albini': 'SOL23521',
               'Katherine England/Abini': 'SOL23521', 'Justin Locke': 'SOL4967', 'Brian Bennett': 'SOL3055',
               'Carter Roedell': 'SOL23345', 'Fernanda Hernandez': 'SOL26788', 'Fern Hernandez': 'SOL26788',
               'Alton Major': 'SOL4809', 'Thuy Pham': 'SOL25688', 'Julianne Martinez': 'SOL22766',
               'Quenton Stroud': 'SOL27228', 'Sadie Oliver': 'SOL26834', 'Valeria Rebollar': 'SOL24218',
               'Sergio Espinoza': 'SOL23542', 'Olivia Larimer': 'SOL5463', 'Grayson Corbin': 'SOL1604',
               'Deonte Keller': 'SOL27498', 'Rayven Alexander': 'SOL24125', 'Deeandra Castillo': 'SOL5495',
               'Kenan Williams': 'SOL27567'}


class ConfirmationSheet:

    def __init__(self, file_name):
        self.file_name = file_name
        self.excel_dataframe = 0

    def convert_to_csv(self):
        xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\Downloads\\" + self.file_name)
        df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
        df.to_csv('file.csv')
        self.file_name = 'file.csv'

    def count_pids(self):
        number_of_pids = 0
        with open(self.file_name) as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                number_of_pids += 1
            return number_of_pids


class Row:

    def __init__(self, pid, c, r, x, u, t, completed):
        self.pid = pid
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
        pyautogui.doubleClick(m2['prospect_id'])
        keyboard.send('ctrl + c')
        copied_text = clipboard.paste()
        for i in range(3):
            if copied_text != self.pid:
                time.sleep(0.3)
                pyautogui.doubleClick(m2['prospect_id'])
                keyboard.send('ctrl + c')
                copied_text = clipboard.paste()
        if copied_text != self.pid:
            cf.pause('Is the pid correct?')
            return
        if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\company.png',
                                          region=(514, 245, 889, 566)) is not None:
            return
        pyautogui.click(m2['company'])
        keyboard.send('ctrl + z')
        time.sleep(1)
        keyboard.send('ctrl + c')
        copied_text = clipboard.paste()
        if 'pid' in copied_text.lower():
            cf.pause('Is the pid correct?')
            return

    def select_tour(self):
        x, y = m2['title']
        current_date = cf.get_current_date()
        df, pretty_df = cf.create_accommodations_dataframe()
        print(tabulate(pretty_df, headers='keys', tablefmt='psql'))
        # Returns the top tour that is Showed, not an Audition, and at most a week before the date we entered.
        # tour_number is the index of the correct tour. Ex: 1 if the second tour is the correct one.
        if self.c in ['x', 'X']:
            try:
                tour_number = df[((df.Tour_Status == 'Showed') | (df.Tour_Status == 'Confirmed') |
                                  (df.Tour_Status == 'No_Show') | (df.Tour_Status == 'On_Tour')) &
                                 (df.Tour_Type != 'Audition') &
                                 ((df.Date - current_date) >= datetime.timedelta(days=-1)) &
                                 ((df.Date - current_date) <= datetime.timedelta(days=14))].index[0]
            except IndexError:
                print('Couldn\'t find correct tour')
                tour_number = df[(df.Tour_Type != 'Audition')].index[0]
        elif self.r in ['x', 'X']:
            try:
                tour_number = df[((df.Tour_Status == 'Rescheduled') &
                                  ((df.Date - current_date) >= datetime.timedelta(days=0))) |
                                 ((df.Tour_Type == 'Open_Reservation') &
                                  (df.Date == datetime.datetime.strptime('1/1/1900', "%m/%d/%Y"))) &
                                 (df.Tour_Type != 'Audition')].index[0]
            except IndexError:
                print('Couldn\'t find correct tour')
                tour_number = df[(df.Tour_Type != 'Audition')].index[0]
        elif self.x in ['x', 'X']:
            try:
                tour_number = df[(df.Tour_Status == 'Canceled') & (df.Tour_Type != 'Audition')].index[0]
            except IndexError:
                print('Couldn\'t find correct tour')
                tour_number = df[(df.Tour_Type != 'Audition')].index[0]
        elif self.u in ['x', 'X']:
            try:
                tour_number = df[(df.Tour_Type == 'Minivac') &
                                 ((df.Date - current_date) >= datetime.timedelta(days=-1))].index[0]
            except IndexError:
                print('Couldn\'t find correct tour')
                tour_number = df[(df.Tour_Type != 'Audition')].index[0]
        elif self.t in ['x', 'X']:
            try:
                tour_number = df[(df.Tour_Type == 'Day_Drive') &
                                 ((df.Date - current_date) >= datetime.timedelta(days=-1))].index[0]
            except IndexError:
                print('Couldn\'t find correct tour')
                tour_number = df[(df.Tour_Type != 'Audition')].index[0]
        else:
            print('Couldn\'t find correct tour')
            tour_number = df[(df.Tour_Type != 'Audition')].index[0]
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


class Tour:

    def __init__(self):
        self.campaign = None
        self.tour_type = None
        self.tour_status = None
        self.tour_date = None
        self.tour_location = None
        self.wave = None

    def gather_info(self):
        sc.get_m3_coordinates()
        x, y = m3['title']
        tour_types_dict = cf.read_pickle_file('m3_tour_type.p')
        self.tour_type = tour_types_dict[cf.take_screenshot(x + 36, y + 143, 89, 12)]
        self.tour_status = sc.m3_tour_status[cf.take_screenshot(x + 37, y + 170, 94, 11)]
        month = cf.take_screenshot(x + 37, y + 196, 13, 10)
        day = cf.take_screenshot(x + 52, y + 196, 15, 10)
        year = cf.take_screenshot(x + 68, y + 196, 27, 10)
        self.tour_date = cf.turn_screenshots_into_date(month, day, year)
        try:
            self.tour_date = datetime.datetime.strptime(self.tour_date, "%m/%d/%Y")
        except ValueError:
            month = cf.take_screenshot(x + 40, y + 196, 13, 10)
            day = cf.take_screenshot(x + 55, y + 196, 15, 10)
            year = cf.take_screenshot(x + 71, y + 196, 27, 10)
            self.tour_date = cf.turn_screenshots_into_date(month, day, year)
            self.tour_date = datetime.datetime.strptime(self.tour_date, "%m/%d/%Y")
        return self.tour_type, self.tour_status, self.tour_date


confirmation_sheet = ConfirmationSheet('3.xlsx')
confirmation_sheet.convert_to_csv()
number_of_pids = confirmation_sheet.count_pids()
with open('file.csv') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if row['Sol'] != '':
            sol = row['Sol']
        row = Row(row['PID'].replace('.0', ''), row['conf'], row['rxl'], row['cxl'], row['ug'], row['tav'], row['Completed'])
        if row.completed == 'x':
            continue
        row.search_pid()
        row.double_check_pid()
        row.select_tour()
        tour = Tour()
        tour.gather_info()
        print(tour.tour_type)
        print(tour.tour_status)
        print(tour.tour_date)

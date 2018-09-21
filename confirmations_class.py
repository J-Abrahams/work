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

    def get_sol(self):
        with open(self.file_name) as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if row['Sol'] != '':
                    try:
                        sol = sol_numbers[row['Sol']]
                    except TypeError:
                        print('Unrecognized Name - ' + sol)
                        sol = 'SOL' + input('Type Sol number (just numbers):')


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


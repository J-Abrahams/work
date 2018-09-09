import keyboard
import pyautogui
import time
import sys
import datetime
import clipboard
import mss
import mss.tools
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8
import pandas as pd
import csv


def take_screenshot(x, y, width, height, save_file=False):
    with mss.mss() as sct:
        monitor = {'top': y, 'left': x, 'width': width, 'height': height}
        im = sct.grab(monitor)
        screenshot = str(mss.tools.to_png(im.rgb, im.size))
        if save_file:
            now = datetime.datetime.now()
            output = now.strftime("%m-%d-%H-%M-%S.png".format(**monitor))
            mss.tools.to_png(im.rgb, im.size, output=output)
        return screenshot


class ExcelSheet:

    def __init__(self, file_name):
        self.file_name = file_name
        self.number_of_pids = 0
        self.excel_dataframe = 0
        self.progress = 0
        self.pid = 0
        self.c = 0
        self.x = 0
        self.r = 0

    def convert_excel_to_csv(self):
        xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\Downloads\\" + self.file_name)
        df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
        self.excel_dataframe = df
        print(type(self.excel_dataframe))
        df.to_csv('file.csv')
        self.file_name = 'file.csv'

    def assign_variable(self):
        pid = self.excel_dataframe['PID'][self.progress]
        self.c = self.excel_dataframe['conf'][self.progress]
        self.x = self.excel_dataframe['cxl'][self.progress]
        self.r = self.excel_dataframe['rxl'][self.progress]
        u = self.excel_dataframe['ug'][self.progress]
        t = self.excel_dataframe['tav'][self.progress]
        if self.excel_dataframe['conf'][self.progress] == 'NAN':
            sol = self.excel_dataframe['Sol'][self.progress]
        completed = self.excel_dataframe['Completed'][self.progress]

    def count_number_of_pids(self):
        with open(self.file_name) as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                self.number_of_pids += 1

    def print_progress(self):
        excel.count_number_of_pids()
        percentage = round((self.progress - 1) * 100 / self.number_of_pids, 2)
        print('{} / {} - {} - {}'.format(str(self.progress), str(self.number_of_pids), str(percentage) + '%', str(pid)))

    def get_deposit_info(self):
        d = []
        with open(self.file_name) as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                pid = row['PID'].replace('.0', '')
                price = row['price']
                date = row['date']
                date = datetime.datetime.strptime(date, '%Y-%m-%d').strftime('%m/%d')
                cash = row['cash']
                d.append({'PID': pid, 'Price': price, 'Date': date, 'Cash': cash})
        df = pd.DataFrame(d)
        self.excel_dataframe = df
        print(self.excel_dataframe)


class Deposit:

    def __init__(self, pid, amount, date, cash_or_cc):
        self.pid = pid
        self.amount = amount
        self.date = datetime.datetime.strptime(date, '%Y-%m-%d').strftime('%m/%d')
        self.cash_or_cc = cash_or_cc

    def get_descriptive_name(self):
        """Return a neatly formatted descriptive name."""
        long_name = str(self.pid) + ' ' + str(self.amount) + ' ' + str(self.date) + ' ' + str(self.cash_or_cc)
        return long_name

    def change_description(self):
        sc.get_m3_coordinates()
        price = 0
        old_title = 'old'
        while self.amount != price:
            pyautogui.click(m3['tour_packages'])
            pyautogui.click(m3['deposit_1'])
            pyautogui.click(m3['change_deposit'])
            sc.get_m6_coordinates()
            x, y = m6['title']
            try:
                price = sc.deposit_item_amount[take_screenshot(x + 168, y + 187, 51, 11)]
            except KeyError:
                price = 0
            take_screenshot(x + 168, y + 187, 51, 11)
            with mss.mss() as sct:
                x, y = m6['title']
                monitor = {'top': y + 187, 'left': x + 168, 'width': 51, 'height': 11}
                im = sct.grab(monitor)
                try:
                    amount = sc.screenshot_dict[str(mss.tools.to_png(im.rgb, im.size))]
                except KeyError:
                    amount = 0
            if amount != price:
                pyautogui.click(m6['ok'])
                y_2 += 13
        while 'ref' not in old_title.lower():
            pyautogui.click(m6['description'])
            keyboard.send('ctrl + z')
            keyboard.send('ctrl + c')
            r = Tk()
            old_title = r.selection_get(selection="CLIPBOARD")
        new_title = old_title.replace("able", "ed")
        new_title = new_title.replace("ABLE", "ED")
        new_title = new_title.replace(" /", "/")
        new_title = new_title.replace("/ ", "/")
        keyboard.write(new_title)


excel = ExcelSheet('3.xlsx')
excel.convert_excel_to_csv()
excel.count_number_of_pids()
excel.assign_variable()
print(type(excel.c))
print(excel.r)
excel.get_deposit_info()
print(excel.excel_dataframe())

deposit = Deposit(1400806, 50, '2018-08-15', 'cc')
print(deposit.get_descriptive_name())
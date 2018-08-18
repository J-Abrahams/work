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


class ExcelSheet:

    def __init__(self, file_name):
        self.file_name = file_name
        self.number_of_pids = 0
        self.excel_dataframe = 0

    def convert_excel_to_csv(self):
        xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\Downloads\\" + self.file_name)
        df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
        df.to_csv('file.csv')
        self.file_name = 'file.csv'

    def count_number_of_pids(self):
        with open(self.file_name) as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                self.number_of_pids += 1

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
        amount = 0
        old_title = 'old'
        x_2, y_2 = m3['deposit_1']
        while amount != price:
            pyautogui.click(m3['tour_packages'])
            pyautogui.click(x_2, y_2)
            pyautogui.click(m3['change_deposit'])
            sc.get_m6_coordinates()
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


excel = ExcelSheet('deposit_pids.xlsx')
excel.convert_excel_to_csv()
excel.count_number_of_pids()
print(excel.number_of_pids)
excel.get_deposit_info()
excel.excel_dataframe()

deposit = Deposit(1400806, 50, '2018-08-15', 'cc')
print(deposit.get_descriptive_name())

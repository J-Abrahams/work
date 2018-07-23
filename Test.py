import csv
import time
from tkinter import Tk
import keyboard
import mss
import mss.tools
import pandas as pd
import pyautogui
import pyperclip
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8


def search_pid(pid_number):
    sc.get_m1_coordinates()
    pyautogui.doubleClick(m1['search'])
    keyboard.write(pid_number)
    pyautogui.click(m1['find_now'])
    pyautogui.click(m1['change'])


def enter_ph_number(number):
    sc.get_m2_coordinates()
    pyautogui.doubleClick(m2['phone2'])
    keyboard.write(number)
    keyboard.send('tab')
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                           region=(514, 245, 889, 566))
    while image is None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_date.png',
                                               region=(514, 245, 889, 566))
    x, y = image
    pyautogui.click(x - 20, y + 425)


with open('phone.csv') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        pids = row['PID'].replace('.0', '')
        phone_number = row['phone']
        search_pid(pids)
        enter_ph_number(phone_number)
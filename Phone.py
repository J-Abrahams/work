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
import importlib


def search_pid(pid_number):
    sc.get_m1_coordinates()
    pyautogui.doubleClick(m1['search'])
    keyboard.write(pid_number)
    pyautogui.click(m1['find_now'])
    pyautogui.click(m1['change'])


def enter_phone_number(number):
    sc.get_m2_coordinates()
    screen_shot = None
    pyautogui.doubleClick(m2['phone2'])
    keyboard.write(number)
    keyboard.send('tab')
    pyautogui.click(m2['ok'])
    while screen_shot != sc.phone_error and screen_shot != sc.phone_no_error:
        with mss.mss() as sct:
            monitor = {'top': 306, 'left': 722, 'width': 38, 'height': 16}
            im = sct.grab(monitor)
            screen_shot = str(mss.tools.to_png(im.rgb, im.size))
    if screen_shot == sc.phone_error:
        pyautogui.click(740, 313)
        pyautogui.click(m2['ok'])
        return "Error"
    elif screen_shot == sc.phone_no_error:
        return "Good"


with open('Phone_Errors.txt', 'w') as erase:
    erase.write('')
with open('phone.csv') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        pids = row['PID'].replace('.0', '')
        phone_1 = row['phone_1']
        phone_2 = row['phone_2']
        if phone_1 != phone_2:
            search_pid(pids)
            status = enter_phone_number(phone_2)
            if status == "Error":
                with open('Phone_Errors.txt', 'a') as out:
                    out.write('{} {} {}\n'.format(pids, phone_1, phone_2))

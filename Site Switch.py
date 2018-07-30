import importlib
import re
import csv
import time
from tkinter import Tk
import keyboard
import mss
import mss.tools
import pandas as pd
import pyautogui
import clipboard
import pyperclip
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8
import datetime


def get_bad_tour_info():
    campaign = get_campaign()
    tour_type = get_tour_type()


def get_campaign():
    sc.get_m3_coordinates()
    pyautogui.click(m3['campaign'])
    sc.get_m4_coordinates()
    pyautogui.doubleClick(m4['campaign'])
    keyboard.send('ctrl + c')
    campaign = clipboard.paste()
    pyautogui.click(m4[''])
    return campaign


def get_tour_type():
    sc.get_m3_coordinates()
    with mss.mss() as sct:
        x, y = m3['title']
        monitor = {'top': y + 170, 'left': x + 37, 'width': 79, 'height': 11}
        im = sct.grab(monitor)
        tour_type = str(mss.tools.to_png(im.rgb, im.size))
        clipboard.copy(tour_type)

"""bad_pid = input('Enter the bad PID:')
good_pid = input('Input the good PID:')
confirmation_message  = input('The bad tour should be open and should be window 1. The good tour should be ready in '
                              'window 2.')"""
get_tour_type()
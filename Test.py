import numpy as np
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
import pyperclip
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8
import datetime


dict = {}
d = []
sc.get_m2_coordinates()
x, y = m2['title']
for i in range(5):
    with mss.mss() as sct:
        monitor = {'top': y + 63, 'left': x + 400, 'width': 75, 'height': 11}
        im = sct.grab(monitor)
        try:
            screenshot = sc.m2_tour_types[str(mss.tools.to_png(im.rgb, im.size))]
        except KeyError:
            print(i)
            print(str(mss.tools.to_png(im.rgb, im.size)))
            screenshot = None
        monitor = {'top': y + 63, 'left': x + 483, 'width': 73, 'height': 11}
        im = sct.grab(monitor)
        try:
            screenshot_2 = sc.m2_tour_status[str(mss.tools.to_png(im.rgb, im.size))]
        except KeyError:
            print(i)
            print(str(mss.tools.to_png(im.rgb, im.size)))
            screenshot_2 = None
        y += 13
        try:
            d.append({'Tour Type': screenshot, 'Tour Status': screenshot_2})
        except NameError:
            pass

df = pd.DataFrame(d)
print(df)
for i, row in df.iterrows():
    print(row['Tour Type'])
    if row['Tour Status'] == 'day_drive' and row['']:
        index = i

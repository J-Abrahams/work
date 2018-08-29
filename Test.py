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

d = []
sc.get_m2_coordinates()
x, y = m2['title']
for i in range(5):
    with mss.mss() as sct:
        monitor = {'top': y + 63, 'left': x + 402, 'width': 14, 'height': 10}
        im = sct.grab(monitor)
        try:
            screenshot = sc.m2_tour_types[str(mss.tools.to_png(im.rgb, im.size))]
        except KeyError:
            print(i)
            print(str(mss.tools.to_png(im.rgb, im.size)))
            screenshot = None
        monitor = {'top': y + 63, 'left': x + 484, 'width': 14, 'height': 10}
        im = sct.grab(monitor)
        try:
            screenshot_2 = sc.m2_tour_status[str(mss.tools.to_png(im.rgb, im.size))]
        except KeyError:
            print(i)
            print(str(mss.tools.to_png(im.rgb, im.size)))
            screenshot_2 = None
        y += 13
        try:
            d.append({'Tour_Type': screenshot, 'Tour_Status': screenshot_2})
        except NameError:
            pass

x, y = m2['title']
df = pd.DataFrame(d)
cols = df.columns.tolist()
cols = cols[-1:] + cols[:-1]
df = df[cols]
print(df)
tour_number = df[df.Tour_Status == 'Showed'].index[0]
pyautogui.doubleClick(x + 469, y + 67 + 13 * tour_number)

for i, row in df.iterrows():
    print(row['Tour_Type'])
    if row['Tour_Status'] == 'day_drive' and row['']:
        index = i
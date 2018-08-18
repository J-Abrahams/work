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

for H in range(0,len(a_list)):
    if a_list[H] > list4[0]:
        list5 = [number_list[i]]
        if function(list1,list5) == list1[1]:
            if function(list2,list5)== list2[1]:
                if function(list3,list5)== list3[1]:
                    if function(list4,list5)== list4[1]:
                        list5.append(input('some input from the user'))
                        other_function(list5)
                        if list5[1]== 40:
                            print ('something something')
                            break out of EVERY loop
                         else:
                            for H in range(0,len(a_list)):
                                if a_list[H] > list5[0]:
                                    list6 = [number_list[i]]
                                    if function(list1,list6) == list1[1]:
                                        if function(list2,list6)== list2[1]:
                                            if function(list3,list6)== list3[1]:
                                               if function(list4,list6)== list4[1]:
                                                  if function(list5,list6)== list5[1]:
                                                     list6.append(input('some input from theuser'))
                                                     other_function(list6)
                                                         if list6[1]== 40:
                                                             print ('something something')
                                                                 break out of EVERY loop
                                                         else:
                                                            etc. (one extra comparison every time)


for H in range(0,len(a_list)):
    if a_list[H] > lst[3][0]:
        lstA = [number_list[i]]
        if function(lst[0],lstA) == lst[0][1]:
            if function(lst[1],lstA)== lst[1][1]:
                if function(lst[2],lstA)== lst[2][1]:
                    if function(lst[3],lstA)== lst[3][1]:
                        lstA.append(input('some input from the user'))
                        other_function(lstA)
                        if lstA[1]== 40:
                            print ('something something')
                            break out of EVERY loop
                         else:
                            for H in range(0,len(a_list)):
                                if a_list[H] > lstA[0]:
                                    lstB = [number_list[i]]
                                    if function(lst[0],lstB) == lst[0][1]:
                                        if function(lst[1],lstB)== lst[1][1]:
                                            if function(lst[2],lstB)== lst[2][1]:
                                               if function(lst[3],lstB)== lst[3][1]:
                                                  if function(lstA,lstB)== lstA[1]:
                                                     lstB.append(input('some input from theuser'))
                                                     other_function(lstB)
                                                         if lstB[1]== 40:
                                                             print ('something something')
                                                                 break out of EVERY loop
                                                         else:
                                                            etc. (one extra comparison every time)
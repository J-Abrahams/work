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

sc.get_m3_coordinates()
with mss.mss() as sct:
    sct.shot(output='1')
pyautogui.click(m3['user_fields'])
pyautogui.click(m3['tour_packages'])
with mss.mss() as sct:
    sct.shot(output='2')
pyautogui.click(m3['notes'])
pyautogui.click(m3['premiums'])
with mss.mss() as sct:
    sct.shot(output='3')

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
import core_functions as cf


def take_screenshot(x, y, width, height, name, save_file=False):
    with mss.mss() as sct:
        monitor = {'top': y, 'left': x, 'width': width, 'height': height}
        im = sct.grab(monitor)
        screenshot = str(mss.tools.to_png(im.rgb, im.size))
        if save_file:
            output = name
            mss.tools.to_png(im.rgb, im.size, output=output)
        return screenshot


sc.get_m3_coordinates()
take_screenshot(498, 254, 1130, 723, '1.png', True)
pyautogui.click(m3['user_fields'])
pyautogui.click(m3['tour_packages'])
take_screenshot(498, 254, 1130, 723, '2.png', True)
pyautogui.click(m3['notes'])
pyautogui.click(m3['premiums'])
take_screenshot(498, 254, 1130, 723, '3.png', True)

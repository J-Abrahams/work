import pyautogui
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10
import keyboard
import mss
import mss.tools
import datetime
import pickle
import sys


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


def read_pickle_file(file_name):
    with open('text_files\\' + file_name, 'rb') as file:
        return pickle.load(file)


def enter_deposit():
    sc.get_m3_coordinates()
    deposit_options_dictionary = read_pickle_file('deposit_options.p')
    pyautogui.click(m3['tour_packages'])
    pyautogui.click(m3['insert_deposit'])
    sites_dictionary = read_pickle_file('sites.p')
    site = take_screenshot(1517, 1036, 146, 17)
    sc.get_m6_coordinates()
    pyautogui.click(m3)
    keyboard.write('SOL/Refundable Deposit')
    pyautogui.click(m6['insert'])
    sc.get_m8_coordinates()
    pyautogui.click(m8['transaction_code'])
    pyautogui.click(m8['transaction_code_scroll_bar'])
    x, y = m8['title']
    if sites_dictionary[site] == 'A3':
        for i in range(9):
            deposit_option = take_screenshot(x + 32, y + 91, 135, 1)
            if deposit_options_dictionary[deposit_option] == 'sol credit payment':
                pyautogui.click(x + 75, y + 91)
                break
            elif i == 8:
                sys.exit("Couldn't find correct option.")
            else:
                y += 13


enter_deposit()

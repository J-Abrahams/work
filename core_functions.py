import mss.tools
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10, m11
import pyautogui
import keyboard
import pickle
import tabulate
import clipboard
import pandas as pd
import time
import datetime
import csv
import random
import cv2
import numpy as np


def search_pid(pid_number):
    sc.get_m1_coordinates()
    pyautogui.doubleClick(m1['search'])
    keyboard.write(pid_number)
    pyautogui.click(m1['find_now'])
    pyautogui.click(m1['change'])


def double_check_pid(pid_number):
    sc.get_m2_coordinates()
    pyautogui.doubleClick(m2['prospect_id'])
    keyboard.send('ctrl + c')
    copied_text = clipboard.paste()
    for i in range(3):
        if copied_text != pid_number:
            time.sleep(0.3)
            pyautogui.doubleClick(m2['prospect_id'])
            keyboard.send('ctrl + c')
            copied_text = clipboard.paste()
    if copied_text != pid_number:
        pause('Is the pid correct?')
        return
    if pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\company.png',
                                      region=(514, 245, 889, 566)) is not None:
        return
    pyautogui.click(m2['company'])
    keyboard.send('ctrl + z')
    time.sleep(1)
    keyboard.send('ctrl + c')
    copied_text = clipboard.paste()
    if 'pid' in copied_text.lower():
        pause('Is the pid correct?')
        return


def create_accommodations_dataframe():
    """
    Takes screenshots of the tours. Turns the screenshots into a list of dictionaries 'd'. Turns 'd' into a dataframe.
    d is a list of dictionaries such as [{''Date': '7/06/18', 'Tour_Type': 'Audition', 'Tour_Status': 'Showed'},
    {'Date': '7/06/18', 'Tour_Type': 'minivac', 'Tour_Status': 'Showed'}]
    :return:
    """
    sc.get_m2_coordinates()
    d = []
    pretty_d = []
    x, y = m2['title']
    for i in range(8):

        # Take screenshots of dates
        month = take_screenshot(x + 327, y + 63, 13, 10)
        day = take_screenshot(x + 342, y + 63, 15, 10)
        year = take_screenshot(x + 358, y + 63, 27, 10)

        # Turn the screenshots into a datetime object
        tour_date = turn_screenshots_into_date(month, day, year)

        # Take screenshots of the tour type and tour status
        tour_type = take_screenshot(x + 402, y + 63, 14, 10)
        tour_status = take_screenshot(x + 484, y + 63, 14, 10)

        # TODO Use pickle file here instead of dictionary in screenshot_data.
        # Turn screenshot of tour type into string using the dictionary m2_tour_types
        try:
            tour_type = sc.m2_tour_types[tour_type]
        except KeyError:
            print('Unrecognized tour type')
            print(tour_type)
            tour_type = None

        # Turn screenshot of tour status into string using the pickle file m2_tour_status
        try:
            tour_status_dict = read_pickle_file('m2_tour_status.p')
            tour_status = tour_status_dict[tour_status]
        except KeyError:
            print('Unrecognized tour status')
            print(tour_status)
            tour_status = None
        y += 13
        # Turns the screenshots into dictionaries.
        if tour_date != 'Nothing':
            try:
                # pretty_d uses strings instead of datetime objects because strings look better.
                pretty_d.append({'Date': tour_date, 'Tour_Type': tour_type, 'Tour_Status': tour_status})

                # Turn strings into datetime objects and creates the dictionaries for the actual dataframe.
                tour_date = pd.to_datetime(tour_date)
                d.append({'Date': tour_date, 'Tour_Type': tour_type, 'Tour_Status': tour_status})
            except NameError:
                pass
        elif tour_status == 'Error':
            pretty_d.append({'Date': '', 'Tour_Type': '', 'Tour_Status': 'Error'})
            d.append({'Date': '', 'Tour_Type': '', 'Tour_Status': 'Error'})

    # Turns d into a dataframe and then reorders the columns in the dataframe.
    df = pd.DataFrame(d)
    df = df[['Date', 'Tour_Type', 'Tour_Status']]

    # Turns pretty_d into a dataframe and then reorders the columns in the dataframe.
    pretty_df = pd.DataFrame(pretty_d)
    pretty_df = pretty_df[['Date', 'Tour_Type', 'Tour_Status']]

    # Prints the pretty dataframe, returns the actual dataframe.
    return df, pretty_df


def take_screenshot(x, y, width, height, save_file=False):
    with mss.mss() as sct:
        sct.compression_level = 3
        monitor = {'top': y, 'left': x, 'width': width, 'height': height}
        im = sct.grab(monitor)
        screenshot = str(mss.tools.to_png(im.rgb, im.size))
        if save_file:
            now = datetime.datetime.now()
            date = now.strftime("%m-%d-%H-%M-%S".format(**monitor))
            output = date + str(random.randrange(0, 1000, 1)) + '.png'
            mss.tools.to_png(im.rgb, im.size, output=output)
        return screenshot


def take_screenshot_change_color(x, y, width, height, save_file=False):
    with mss.mss() as sct:
        monitor = {'top': y, 'left': x, 'width': width, 'height': height}
        # monitor = {'top': 139, 'left': 284, 'width': 80, 'height': 11}
        img = np.array(sct.grab(monitor))

    height, width, channels = img.shape

    blue = [107, 36, 8, 255]
    white = [255, 255, 255, 255]
    black = [0, 0, 0, 255]

    if np.any(img[:, 0] == 107):
        for x in range(0, width):
            for y in range(0, height):
                channels_xy = img[y, x]
                if all(channels_xy == white):
                    img[y, x] = black
                elif all(channels_xy == blue):
                    img[y, x] = white
    if save_file:
        cv2.imshow('img', img)
    screenshot = cv2.imencode('.png', img)[1]
    return str(screenshot.tobytes())


def get_current_date():
    now = datetime.datetime.now()
    now = now.strftime("%m/%d/%y")
    return datetime.datetime.strptime(now, "%m/%d/%y")


def turn_screenshots_into_date(month_screenshot, day_screenshot, year_screenshot):
    f = open('text_files\\dates.p', 'rb')
    date_dictionary = pickle.load(f)
    f.close()
    try:
        month_screenshot = date_dictionary[month_screenshot]
        day_screenshot = date_dictionary[day_screenshot]
        year_screenshot = date_dictionary[year_screenshot]
        if month_screenshot not in ['Nothing', 'Error']:
            tour_date = ('{}/{}/{}'.format(month_screenshot, day_screenshot, year_screenshot))
            return tour_date
            # datetime.datetime.strptime(tour_date, "%m/%d/%Y")
        elif month_screenshot == 'Error':
            return 'Error'
        elif month_screenshot == 'Nothing':
            return 'Nothing'
    except KeyError:
        return 'Nothing'


def print_colored_text(text, color):
    if color == 'red':
        print('{}{}{}'.format(u"\u001b[31m", text, u"\u001b[0m"))
    elif color == 'green':
        print('{}{}{}'.format(u"\u001b[32m", text, u"\u001b[0m"))


def remove_duplicate_pickle_keys():
    """
Removes duplicate dictionary keys from the premiums.p file.
    """
    global premiums
    global premium_dict
    premiums = {}
    for key, value in premium_dict.items():
        if key not in premiums.keys():
            premiums[key] = value
    f = open('text_files\\premiums.p', 'wb')
    pickle.dump(premiums, f)
    f.close()


def read_pickle_file(file_name):
    with open('text_files\\' + file_name, 'rb') as file:
        return pickle.load(file)


def pause(message):
    print(message)
    while pyautogui.position() != (0, 1079):
        pass


def open_excel_sheet_as_csv(filename):
    xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\Downloads\\" + filename)
    df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
    df.to_csv('file.csv')
    file = open('file.csv')
    excel_sheet = csv.DictReader(file)
    file.close()
    return excel_sheet
    with open('file.csv') as csvfile:
        return csv.DictReader(csvfile)


def count_number_of_pids():
    number_of_pids = 0
    with open('file.csv') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            number_of_pids += 1
    return number_of_pids

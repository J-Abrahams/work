from __future__ import print_function
import csv
import keyboard
import mss
import mss.tools
import pyautogui
import screenshot_data as sc
from screenshot_data import m1, m2
import datetime
import core_functions as cf
from oauth2client.service_account import ServiceAccountCredentials
import gspread


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


scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('Phone-6ad41718c799.json', scope)
client = gspread.authorize(creds)
sheet = client.open("Phone Errors").sheet1

cf.pause('Minimize Pycharm')
now = datetime.datetime.now()
now_str = now.strftime("%m/%d/%Y")
# with open('text_files\\phones\\Phone_Errors.txt', 'a') as erase:
#     erase.write('\n{}\n'.format(now_str))
with open('text_files\\phones\\phone.csv', 'r') as csvfile:
    reader = csv.DictReader(csvfile)
    number_of_phone_numbers = 0
    for row in reader:
        number_of_phone_numbers += 1
with open('text_files\\phones\\phone.csv') as csvfile:
    reader = csv.DictReader(csvfile)
    duplicate_phone_numbers, progress, errors = 0, 0, 0
    for row in reader:
        pids = row['PID'].replace('.0', '')
        phone_1 = row['phone_1']
        phone_2 = row['phone_2']
        if phone_1 != phone_2 and len(phone_1) == len(phone_2):
            search_pid(pids)
            status = enter_phone_number(phone_2)
            if status == "Error":
                if errors == 0:
                    sheet.append_row([now_str, pids, phone_1, phone_2])
                else:
                    sheet.append_row(['', pids, phone_1, phone_2])
                errors += 1
                # with open('text_files\\phones\\Phone_Errors.txt', 'a') as out:
                #     out.write('{} {} {}\n'.format(pids, phone_1, phone_2))
        elif len(phone_1) != len(phone_2):
            if errors == 0:
                sheet.append_row([now_str, pids, phone_1, phone_2])
            else:
                sheet.append_row(['', pids, phone_1, phone_2])
            errors += 1
        else:
            duplicate_phone_numbers += 1
        progress += 1
        print(str(progress) + '/' + str(number_of_phone_numbers))
        print(str(duplicate_phone_numbers) + ' duplicate numbers')
        print(str(errors) + ' errors')

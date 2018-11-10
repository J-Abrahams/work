import csv
import time
from tkinter import Tk
import keyboard
import mss
import mss.tools
import pandas as pd
import pyautogui
import clipboard
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10, m11, m12, m13
import logging
import datetime
from tabulate import tabulate
import sys
import pickle
import re
import core_functions as cf
import sqlite3
from oauth2client.service_account import ServiceAccountCredentials
import gspread


if __name__ == "__main__":
    source_code = 'so'
    with open('Deposit_Errors.txt', 'a') as out:
        out.write('{}\n'.format(pids))
import sqlite3
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10, m11, m12, m13
import core_functions as cf
import pyautogui
# import importlib
# importlib.reload(cf)

# Connecting to the database file
conn = sqlite3.connect('sqlite.sqlite')
c = conn.cursor()

# Creating a new SQLite table with 1 column
"""c.execute('''CREATE TABLE menu_coordinates (ID INTEGER PRIMARY KEY,
          name TEXT,
          coordinates TEXT)''')"""
"""for key, value in premiums.items():"""
sc.get_m13_coordinates()
x, y = m13['title']
# print(x, y)
# c.execute('INSERT INTO menu_coordinates (name, x, y, description) VALUES (?, ?, ?, ?)', ['m2', 631, 291, 'Changing a Prospect Record'])
c.execute('UPDATE misc SET screenshot=(?) WHERE name=("confirmer")', [cf.take_screenshot_change_color(m13['title'][0] + 26, m13['title'][1] + 73, 45, 15)])
"""for row in c.execute('SELECT * FROM premiums'):
    print(row)"""
"""sc.get_m3_coordinates()
x, y = m3['title']
tour_type = cf.take_screenshot_change_color(x + 36, y + 143, 89, 11)
tour_status = cf.take_screenshot_change_color(x + 36, y + 170, 89, 11)
# c.execute('Insert INTO misc (screenshot, name, description, menu, coordinates) VALUES (?, ?, ?, ?, ?)', [tour_type, 'telesales', 'tour type', 'm3', 'x + 36, y + 143, 89, 11'])
c.execute('Insert INTO misc (screenshot, name, description, menu, coordinates) VALUES (?, ?, ?, ?, ?)', [tour_status, 'booked', 'tour status', 'm3', 'x + 44, y + 16, 9, 27'])
# c.execute("SELECT name FROM premiums WHERE screenshot=?", [screenshot])"""
# rows = c.fetchone()[0]
# print(rows)

# Committing changes and closing the connection to the database file
conn.commit()
conn.close()

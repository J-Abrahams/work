import sqlite3
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10
import core_functions as cf
import pyautogui
# import importlib
# importlib.reload(cf)

# Connecting to the database file
conn = sqlite3.connect('sqlite.sqlite')
c = conn.cursor()

# Creating a new SQLite table with 1 column
"""c.execute('''CREATE TABLE premiums (ID INTEGER PRIMARY KEY,
          name TEXT,
          screenshot TEXT)''')"""
"""for key, value in premiums.items():
    c.execute('INSERT INTO premiums (screenshot, name) VALUES (?, ?)', [key, value])"""
"""c.execute('UPDATE dates SET ID=(?) WHERE name=("1")', [0o1])"""
"""for row in c.execute('SELECT * FROM premiums'):
    print(row)"""
sc.get_m3_coordinates()
x, y = m3['title']
# screenshot = cf.take_screenshot_change_color(284, 139, 80, 11)
screenshot = cf.take_screenshot_change_color(x + 342, y + 60, 80, 11)
c.execute("SELECT name FROM premiums WHERE screenshot=?", [screenshot])
rows = c.fetchone()[0]
print(rows)

# Committing changes and closing the connection to the database file
conn.commit()
conn.close()

import sqlite3
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10

# Connecting to the database file
conn = sqlite3.connect('sqlite.sqlite')
c = conn.cursor()

# Creating a new SQLite table with 1 column
"""c.execute('''CREATE TABLE dates (ID INTEGER PRIMARY KEY,
          name TEXT,
          screenshot TEXT)''')"""
for key, value in sc.date_month.items():
    c.execute('INSERT INTO dates (screenshot, name) VALUES (?, ?)', [key, value])
"""c.execute('UPDATE dates SET ID=(?) WHERE name=("1")', [0o1])"""
for row in c.execute('SELECT * FROM dates'):
    print(row)

# Committing changes and closing the connection to the database file
conn.commit()
conn.close()
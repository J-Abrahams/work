import pandas as pd
xls = pd.ExcelFile("C:\\Users\\Jared.Abrahams\\scan.xlsx")
df = xls.parse(sheet_name="Sheet1", index_col=None, na_values=['NA'])
df.to_csv('file.csv')

import csv

with open('file.csv') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        pids = row['PID']
        pids = pids[:-2]
        print(row['Conf'])
        print(type(pids))
'''
[csv2excel.py]

Copyright (c) 2020 omohikane

This software is released under the MIT License.
http://opensource.org/licenses/mit-license.php
'''

import os
import glob
import pandas as pd
import openpyxl
import pprint
import datetime
from dateutil.relativedelta import relativedelta
from datetime import datetime, date, timedelta
import shutil

def mod_CSV_data_No6_year(changing_year = 1 , date_criteria = '%Y-%m-%d'):
    one_year_ago = lambda year: (datetime.strptime(year, date_criteria) + relativedelta(years=changing_year)).strftime(date_criteria)
    df.iloc[31,2] = one_year_ago(df.iloc[25,2])

def write_list_2d(sheet, l_2d, start_row, start_col):
    for y, row in enumerate(l_2d):
        for x, cell in enumerate(row):
            sheet.cell(row=start_row + y,
                       column=start_col + x,
                       value=l_2d[y][x])

def insert2excel(sheet):
    l_2d = df.values.tolist()
    write_list_2d(sheet, l_2d, 3, 1)

files = glob.glob(r'./csv/*.csv')
for file in files:
    filename, file_extension = os.path.splitext(file)
    out_file = filename + ".xlsx"

    wb = openpyxl.load_workbook('./template.xlsx')
    sheet = wb['sheet1']

    df = pd.read_csv(file, encoding="cp932", skiprows=2, header=None)
    mod_CSV_data_No6_year()

    insert2excel(sheet)
    wb.save(out_file)

    dest = './excel/' + df.iloc[1, 1] + ".xlsx"
    shutil.move(out_file, dest)
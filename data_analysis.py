#!/usr/bin/env python3

"""
Importing data libraries
"""

import pandas as pd
from pandas import ExcelFile
from pandas import ExcelWriter

"""
Defining the location for the excel file along with defining the sheet that we would like to compare, i.e. sheet# 0.
"""

devices= '/Users/touseef.khan/Documents/file.xlsx'
new_week = pd.read_excel(devices, sheet_name = 0)

"""
Loop that opens all the sheets, and then concats the data into one sheet for comparison purposes
"""

xlsx = pd.ExcelFile(devices)
sheets = []
for sheet in xlsx.sheet_names:
    sheets.append(xlsx.parse(sheet))
    new_values = pd.concat(sheets, sort = False)

"""
Compares the data in the sheet, and drops the duplicated values. Then merges any values that might be repeated.
Finally prints the new values in the end on terminal and creates a file with the new values.
"""

new_values.drop_duplicates(subset = "Serial", keep = False, inplace = True)
new_list = pd.merge(New_values, new_week, on = "Values ")
stored = new_list.to_excel('New_values.xlsx', index=None) #stores the result in an excel file along with other information like the other columns
print(new_devices['Values'])

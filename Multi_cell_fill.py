from pickle import FALSE

import pandas as pd
from openpyxl import workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook
from win32com.client import Dispatch

from main import sheet

data = {
    "Asset Name" : ["Asset 1","Asset 2"],
    "Month 1" : [15, 30],
    "Month 2" : [5, 35]
}

df = pd.DataFrame(data)

workbook = Workbook()
sheet = workbook.active

for row in dataframe_to_rows(df, index=False, header=True):
    sheet.append(row)

workbook.save("pandas.xlsx")

x1 = Dispatch("Excel.Application")
x1.Visible = True

wb = x1.Workbooks.open(r'C:\Users\smeduthy\PycharmProjects\Python_to_automate_excel\pandas.xlsx')
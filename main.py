from openpyxl import Workbook
from win32com.client import Dispatch

workbook = Workbook()
sheet = workbook.active

sheet['A1'] = 'hello'
sheet['B1'] = 'Shubash'

workbook.save(filename='mainexcel.xlsx')


x1 = Dispatch("Excel.Application")
x1.Visible = True

wb = x1.Workbooks.open(r'C:\Users\smeduthy\OneDrive - Cisco\Desktop\mainexcel.xlsx')
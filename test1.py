
import win32com.client

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = True
wb = excel.Workbooks.Add()
ws = wb.Worksheets('Sheet1')
ws.Cells(1,1).Value = "Hello"
wb.SaveAs("C:\\Users\\YoungMin\\Documents\\Ubion\\Python Coding\\stocktrading\\test1.xlsx")

import win32com.client
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open("C:\\Users\\YoungMin\\Documents\\Ubion\\Python Coding\\stocktrading\\test1.xlsx")
ws = wb.ActiveSheet
print(ws.Cells(1,1).value)
excel.quit()

import win32com.client
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = True
wb = excel.Workbooks.open("C:\\Users\\YoungMin\\Documents\\Ubion\\Python Coding\\stocktrading\\test1.xlsx")
ws = wb.ActiveSheet
ws.Cells(1,2).Value = 'is'
ws.Range("C1").Value = "good"
ws.Range("C1").interior.colorindex = 10






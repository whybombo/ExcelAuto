from openpyxl import Workbook
from win32com.client import Dispatch

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "hello"
sheet["B1"] = "User"

workbook.save(filename = "hello_User.xlsx")

cell = sheet["A1"]

cell.value = "Greetings"

# print(cell.value)

workbook.save(filename = "hello_User.xlsx")

xl = Dispatch("Excel.Application")
xl.Visible = True

wb = xl.Workbooks.Open(r"C:\Users\Blake Nelson\Desktop\VSC-code\python\AutoExcel-1.0\hello_User.xlsx")
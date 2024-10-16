
# import sys
# sys.path.append(r'C:\Users\<QUOC NHAN>\AppData\Roaming\pyRevit-Master\pyrevitlib')
# sys.path.append(r'C:\Users\<QUOC NHAN>\AppData\Roaming\pyRevit-Master\site-packages')
#import xlwings as xw


import os
import gspread
gs=gspread.service_account("D:\\driveapi\\demoapi-426214-1c72774ca6f5.json")
wb=gs.open_by_key("1MwhMkt7WKHHpjBPt6dTfsEJbATAQb6knagkceKYYVkU")

shts = wb.worksheets()
sht=wb.sheet1 
column_index1 = 1  

column_data1 = sht.col_values(column_index1)

last_value1 = column_data1[-1]
dong_cuoi1=int(len(column_data1))
sht.update_acell(f'b{dong_cuoi1+1}', "456")
from pyrevit import script,forms
from pyrevit import script
from pyrevit.forms import WPFWindow
import os


import clr
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel
from Microsoft.Office.Interop.Excel import*
from System.Runtime.InteropServices import Marshal
#XlListObjectSourceType, Worksheet, Range, XlYesNoGuess,XlReferenceStyle
# from pyrevit import excel
#from xlwings import *
import excel

# #####################################
try:
	# #lien ket voi excel
	# Create a workbook with designated file as template
	#res_path = os.path.join(__commandpath__, r"myfile.xlsx")
	res_path = os.path.join( r"D:\8.bixui\myfile.xlsx")
	#mo excel
	os.startfile(res_path)
	xl_app = excel.initialise()
	res_workbook = xl_app.Workbooks("myfile.xlsx")
	res_sheet = res_workbook.Sheets("Sheet1")
except Exception:
    pass
print(dong_cuoi1)

# selectedCells = res_sheet.Selection
# dc=selectedCells.address
# print(dc)


##################################################################33
#tim dia chi excel
# uidoc = __revit__.ActiveUIDocument
# doc = uidoc.Document

# view = doc.ActiveView

# #print(view.Name)



#####################################

#lien ket voi excel
# xl_app = excel.initialise()

# res_workbook = xl_app.ActiveWorkbook
# res_sheet = res_workbook.Sheets("Sheet1")
# selectedCells = xl_app.Selection
# dc=selectedCells.Address()
# print(dc)


# res_sheet.Range("a3").Select()
# #a=res_sheet.Range("a3").

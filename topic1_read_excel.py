# Reading Excel files
from openpyxl import load_workbook
wb=load_workbook("company_data.xlsx") #open excel file
sheet=wb["Employees"]  # open worksheet
print(sheet.title)    #print sheet name

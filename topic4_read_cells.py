# read  specfic data from excel
from openpyxl import load_workbook
wb=load_workbook("company_data.xlsx")
sheet=wb["Employees"]
print(sheet["A2"].value)
print(sheet["B2"].value)
print(sheet["C2"].value)
print(sheet["D2"].value)
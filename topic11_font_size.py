from openpyxl import load_workbook
from openpyxl.styles import Font

wb = load_workbook("company_data.xlsx")
sheet = wb["Employees"]

# Change header font size
for cell in sheet[1]:
    cell.font = Font(size=16)

wb.save("company_data.xlsx")

print("Font size changed successfully")
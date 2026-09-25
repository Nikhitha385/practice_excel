from openpyxl import load_workbook
from openpyxl.styles import Font

wb = load_workbook("company_data.xlsx")
sheet = wb["Employees"]

# Make Employee ID column bold
for cell in sheet["A"]:
    cell.font = Font(bold=True)

# Make Name column italic
for cell in sheet["B"]:
    cell.font = Font(italic=True)

# Make Department column underlined
for cell in sheet["C"]:
    cell.font = Font(underline="single")

wb.save("company_data.xlsx")

print("Bold, italic and underline applied successfully")
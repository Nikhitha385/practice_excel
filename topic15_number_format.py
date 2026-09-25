from openpyxl import load_workbook

wb = load_workbook("company_data.xlsx")
sheet = wb["Employees"]

# Format salary column
for cell in sheet["D"][1:]:
    cell.number_format = '#,##0'

wb.save("company_data.xlsx")

print("Number format applied successfully")
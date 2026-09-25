from openpyxl import load_workbook

wb = load_workbook("company_data.xlsx")
sheet = wb["Employees"]

# Set column widths
sheet.column_dimensions["A"].width = 15
sheet.column_dimensions["B"].width = 15
sheet.column_dimensions["C"].width = 18
sheet.column_dimensions["D"].width = 15
sheet.column_dimensions["E"].width = 18

# Set header row height
sheet.row_dimensions[1].height = 30

wb.save("company_data.xlsx")

print("Row height and column width updated successfully")
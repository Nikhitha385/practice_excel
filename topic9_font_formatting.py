from openpyxl import load_workbook
from openpyxl.styles import Font

# Open Excel file
wb = load_workbook("company_data.xlsx")

# Open Employees sheet
sheet = wb["Employees"]

# Change the font of the Name column
for cell in sheet["B"]:
    cell.font = Font(name="Arial")

# Save the changes
wb.save("company_data.xlsx")

print("Font formatting completed successfully")
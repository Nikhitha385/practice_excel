#how to create an Excel file using Python.
from openpyxl import Workbook
# Create a new Excel workbook
wb = Workbook()
# Get the active worksheet
sheet = wb.active
# Give the worksheet a name
sheet.title = "Students"
# Save the Excel file
wb.save("students_practice.xlsx")
print("Excel file created successfully")
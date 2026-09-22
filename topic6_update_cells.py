#change existing Excel data using Python.
from openpyxl import load_workbook
wb = load_workbook("company_data.xlsx")
sheet = wb["Employees"]
# Update Rahul's salary
sheet["D2"] = 50000
# Save the changes
wb.save("company_data.xlsx")
print("Cell value updated successfully")
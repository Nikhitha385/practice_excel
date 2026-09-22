from openpyxl import Workbook
# Create a new Excel workbook
wb = Workbook()
# Open the active worksheet
sheet = wb.active
# Give the worksheet a name
sheet.title = "Employees"
# Add headings
sheet.append(["Employee ID", "Name", "Department", "Salary"])
# Add employee data
sheet.append(["E101", "Rahul", "IT", 45000])
sheet.append(["E102", "Priya", "HR", 38000])
sheet.append(["E103", "Arun", "IT", 52000])
sheet.append(["E104", "Sneha", "Finance", 48000])
sheet.append(["E105", "Kiran", "HR", 42000])
# Save the Excel file
wb.save("company_data.xlsx")
print("Excel file created successfully")
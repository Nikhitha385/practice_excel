#write new data into an existing Excel file using Python.
from openpyxl import load_workbook
wb = load_workbook("company_data.xlsx")
sheet = wb["Employees"]
# Write data into cells
sheet["E1"] = "Location"
sheet["E2"] = "Hyderabad"
sheet["E3"] = "Vijayawada"
sheet["E4"]="Banglore"
sheet["E5"]="Eluru"
sheet["E6"]="Tirupathi"
# Save the changes
wb.save("company_data.xlsx")
print("Data written successfully")
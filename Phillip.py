import os
import subprocess
from openpyxl import Workbook

print("Phillip is running from:", os.getcwd())
print("Phillip is starting...")
print("Rendering...")

# Create a new Excel workbook
workbook = Workbook()

# Rename the first sheet
sheet = workbook.active
sheet.title = "Account Log"

# Create columns
sheet["A1"] = "Date"
sheet["B1"] = "Client"
sheet["C1"] = "Product"
sheet["D1"] = "Category"
sheet["E1"] = "Type"
sheet["F1"] = "Amount"

# Add test transactions
transactions = [
    ["2026-09-18", "Office supplies", "Office Supplies", "Expense", 150.00],
    ["2026-09-18", "Customer payment", "Sales", "Income", 1200.00],
    ["2026-09-18", "Shipping", "Shipping", "Expense", 75.00],
]

for transaction in transactions:
    sheet.append(transaction)

# Save the Excel file
output_file = os.path.abspath("accounting.xlsx")

workbook.save(output_file)

print("Accounting spreadsheet created!")
print("Saved to:", output_file)

print("Accounting spreadsheet created!")

subprocess.run(["open", output_file])

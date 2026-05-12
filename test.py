import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Products"

# Headers
ws.append(["Barcode", "Product Name", "Price"])

# Sample products (replace with your real ones)
ws.append(["6001234567890", "Coca Cola 500ml", 25.00])
ws.append(["6009876543210", "Bread 700g", 35.00])
ws.append(["6001111111111", "Milk 1L", 22.00])

wb.save("products.xlsx")
print("products.xlsx created!")
import openpyxl

wb  = openpyxl.Workbook()
sheet = wb.active

sheet.append(["product", "price" ,"Qty"])
sheet.append(["Laptop", 2600 , 5])
sheet.append(["Mouse" , 300, 10])
sheet.append(["Keyboard " ,420, 9])

wb.save("input/sales.xlsx")
print ("sales.xlsx created")
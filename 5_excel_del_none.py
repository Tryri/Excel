import openpyxl
import re

filename = '1.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.active
f = sheet.max_row
k = 0
m = 0
for n in sheet["D"]:
    a = n.value
    # k += 1
    if a == None or a == 'None':
        # print(k)
        sheet.delete_rows(n.row)
        # k = k - 1
        m += 1

wb.save("1.xlsx")
print(m)

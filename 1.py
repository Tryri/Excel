import openpyxl
import re

filename = '1.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.active
f = sheet.max_row
k = 0
m = 0
n1 = 0
n2 = 0
for n in sheet["E"]:
    print(n.row)
    for k in sheet["E"]:
        if n.value == k.value and n.row != k.row and n.value != None:
            n1 = n.row
            n2 = k.row
            m += 1
            print(n.value, k.value)
            # print(n1, n2)
            sheet[f'E{n2}'].value = None
            # sheet[f'C{n2}'].value = None
            # m += 1
wb.save("1.xlsx")
print(m)


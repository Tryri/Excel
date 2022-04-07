import openpyxl
import re

filename = '8.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb['Лист1']
f = sheet.max_row
k = 0
m = 0
n1 = 0
n2 = 0
for n in sheet["A"]:
    for k in sheet["B"]:
        if n.value == k.value and n.value != None:
            n1 = n.row
            n2 = k.row
            m += 1
            print(n.value, k.value)
            # print(n1, n2)
            sheet[f'D{m}'].value = k.value
            sheet[f'A{n1}'].value = None
            sheet[f'B{n2}'].value = None
            # m += 1
wb.save("8.xlsx")
print(m)


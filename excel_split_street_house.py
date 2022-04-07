import openpyxl
import re

filename = '1.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.active
f = sheet.max_row
m = 0
# sheet.insert_cols(2, 1)
for n in sheet["A"]:
    c = sheet[f'A{n.row}'].value
    c = re.split(r',', c)
    sheet[f'A{n.row}'].value = c[0]
    sheet[f'B{n.row}'].value = c[1]
    sheet[f'C{n.row}'].value = c[2]
    print(c)
    m += 1
wb.save("1.xlsx")
print(m)

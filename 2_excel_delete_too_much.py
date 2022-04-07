import openpyxl
import re

filename = '1.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.active
m = 0
f = sheet.max_row
k = 0
for n in sheet['B']:
    c = sheet[f'B{n.row}'].value
    d = sheet[f'A{n.row}'].value
    a = str(c)
    b = re.sub(r' ', '', a)
    b = re.sub(r',$', '', b)
    b = re.split(r',', b)
    if len(b) > 1:
        print(b)
        print(n.row)
        sheet.delete_rows(n.row)
        m += 1

wb.save("1.xlsx")
print(m)
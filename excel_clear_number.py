import openpyxl
import re

filename = '4.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb['Лист1']
f = sheet.max_row
k = 0
for n in range(f):
    k += 1
    c = sheet[f'B{n + 1}'].value
    d = sheet[f'A{n + 1}'].value
    a = str(c)
    b = re.sub(r' ', '', a)
    b = re.sub(r',$', '', b)
    b = re.sub(r'[.]$', '', b)
    b = re.sub(r'\(специализированныйслужебныйжилищныйфонд\)', '', b)
    b = re.sub(r'\(специализированныйманевренныйжилищныйфонд\)', '', b)
    b = re.sub(r'\(специализированныйжилищныйфонддлядетей-сирот\)', '', b)
    b = re.sub(r'долявправеобщейдолевойсобственности', '', b)
    b = re.sub(r'долейвжиломдоме', '', b)
    b = re.sub(r'специализированныйманевренныйжилищныйфонд\)', '', b)
    sheet[f'B{n+1}'].value = b
    # b = re.split(r',', b)
    print(b)
print(k)
wb.save("5.xlsx")
# print(m)

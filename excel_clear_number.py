import openpyxl
import re

filename = '1.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.active
f = sheet.max_row
k = 0
for n in sheet['B']:
    k += 1
    d = sheet[f'B{n.row}'].value
    # d = sheet[f'A{n + 1}'].value
    d = str(d)
    d = re.sub(r' ', '', d)
    d = re.sub(r',$', '', d)
    d = re.sub(r'[.]$', '', d)
    d = re.sub(r'\(специализированныйслужебныйжилищныйфонд\)', '', d)
    d = re.sub(r'\(специализированныйманевренныйжилищныйфонд\)', '', d)
    d = re.sub(r'\(специализированныйжилищныйфонддлядетей-сирот\)', '', d)
    d = re.sub(r'долявправеобщейдолевойсобственности', '', d)
    d = re.sub(r'долейвжиломдоме', '', d)
    d = re.sub(r'специализированныйманевренныйжилищныйфонд\)', '', d)
    sheet[f'D{n.row}'].value = d
    # b = re.split(r',', b)
    print(d)
print(k)
wb.save("1.xlsx")
# print(m)

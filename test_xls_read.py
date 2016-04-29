#import openpyxl

from openpyxl import load_workbook
from openpyxl.cell import get_column_letter
wb = load_workbook(filename = 'Var.xlsx')
sheet = wb['Structs']
#print(sheet_ranges['A1'].value)


#print (wb.get_sheet_names())

col = 'A'
#for row in range(1, 50):
#    cell = col + str(row)
#    print(sheet_ranges[cell].value)

structList = []
row = 2
cell = col + str(row)
while str(sheet[cell].value) != 'None' and row < 1000:
#    print(sheet[cell].value)
    structList.append(str(sheet[cell].value))
    row = row + 1
    cell = col + str(row)

print (structList)

sheet = wb['Vars']

varRow = []
varMatrix = []
row = 2
col = 1
cell = get_column_letter(col) + str(row)
while str(sheet[cell].value) != 'None' and row < 1000:
    for col in range(1, 4):
        cell = get_column_letter(col) + str(row)
        varRow.append(sheet[cell].value)
    row += 1
    cell = get_column_letter(col) + str(row)
    varMatrix.append(varRow)
#    print(varRow)
    varRow = []

print(varMatrix)




#input('press enter any key...')
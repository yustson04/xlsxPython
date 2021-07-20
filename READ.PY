import xlrd

#give location of file
loc = ('D:/_02.Project/_06.imcha/test.xlsx')

#open workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

#for row 0 and column 0
print(sheet.cell_value(0,0))

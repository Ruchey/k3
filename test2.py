import os
import xlsxwriter



pic = r'c:\Users\Александр\Pictures\skylink 2.jpg'
wb = xlsxwriter.Workbook(r'd:\test.xlsx')
ws = wb.add_worksheet('Тест')

cell_format = wb.add_format()

cell_format.set_bottom(1)  # This is optional when using a solid fill.
cell_format.set_bottom_color('green')

ws.write('A1', 'Ray', cell_format)
ws.set_column(1,3,1)

wb.close()
os.startfile(r'd:\test.xlsx')
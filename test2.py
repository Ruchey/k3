from pprint import pprint

import openpyxl


num = 1
proj_rep_path = r'd:\К3\Самара\Самара черновик\1\Reports\Деталировка.xlsx'
wb = openpyxl.load_workbook(proj_rep_path)
ws = wb.active
a =1
wb.close()
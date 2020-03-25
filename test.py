import os
import k3r


fl = r'd:\test.xlsx'
xl = k3r.xl.Doc()
xl.new_sheet('Деталировка')
pic = r'c:\Users\Александр\Pictures\ram.jpg'
# print(xl.get_col_size(10, 6))
print(xl.get_row_size(1, 10))

# xl.paint_cells('A1:E5', fill='cccccc')
# xl.pic_insert(1, 1, path=pic, align='c', valign='c', max_col=5, max_row=5)
xl.paint_cells('A1:E20', fill='cccccc')
xl.pic_insert(1, 1, path=pic, align='l', valign='t', max_col=5, max_row=20)
# xl.pic_insert(1, 6, path=pic, align='l', valign='c', max_col=5, max_row=20)
# xl.pic_insert(1, 12, path=pic, align='l', valign='b', max_col=5, max_row=20)


xl.save(fl)
os.startfile(fl)

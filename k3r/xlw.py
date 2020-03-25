import os
import xlsxwriter


__author__ = 'Виноградов А.Г. г.Белгород  март 2020'


def num_to_col(col_number):
    """Преобразует числовой номер столбца в буквенный"""
    num2col = ""
    while col_number > 0:
        i = int(col_number % 26)
        if i > 0:
            num2col = chr(64 + i) + num2col
        else:
            num2col = "Z" + num2col
            col_number -= 1
        col_number = int(col_number / 26)
    return num2col


def cm_to_inch(cm):
    return cm / 2.54


class Doc:
    """Создание документа Excel
        ExcelDoc - создаёт объект Excel
        newheet - создаёт страницу в книге с заданным именем
        save - сохраняет документ в указанное место
        put_val - добавляет запись в документ
    """

    def __init__(self):
        """Создание объекта Excel, инициализация умолчаний"""
        self.xl = xlsxwriter
        self.wb = self.xl.Workbook(r'd:\test.xlsx')
        self.ws = None

        self.sheet_orient = 'PORTRAIT'
        self.paper_size = 9
        self.right_margin = 0.6
        self.left_margin = 0.6
        self.bottom_margin = 1.6
        self.top_margin = 1.6

        self.center_horizontally = False
        self.font = "Arial Narrow"
        self.font_size = 12
        # self.F_RUB = self.wb.add_format(42)

        self.styles = {}
        self.generate_styles()

    def new_sheet(self, name, tab_color=None):
        """Добавление листа"""

        self.ws = self.wb.add_worksheet(name[:31])
        if self.sheet_orient == 'PORTRAIT':
            self.ws.set_portrait()
        if self.sheet_orient == 'LANDSCAPE':
            self.ws.set_landscape()
        self.ws.set_paper(self.paper_size)
        self.ws.set_margins(left=cm_to_inch(self.left_margin),
                            right=cm_to_inch(self.right_margin),
                            top=cm_to_inch(self.top_margin),
                            bottom=cm_to_inch(self.bottom_margin))
        if self.center_horizontally:
            self.ws.center_horizontally()
        if tab_color:
            self.ws.set_tab_color(tab_color)

    def generate_styles(self):
        """Создание базовых стилей"""

        styles = [
            {'name': 'Заголовок 1', 'bg': 'F2F2F2', 'bc': 'D9D9D9', 'b': True, 'align': 'right',
             'bb': 1},
            {'name': 'Заголовок 2', 'bg': 'D9D9D9', 'bc': 'BFBFBF', 'b': True, 'align': 'right',
             'bb': 1},
            {'name': 'Заголовок 3', 'bg': '8DB4E2', 'bc': '16365C', 'b': True, 'align': 'right',
             'bb': 1},
            {'name': 'Заголовок 4', 'bg': 'B8CCE4', 'bc': '366092', 'b': True, 'align': 'right',
             'bb': 1},
            {'name': 'Заголовок 5', 'bg': 'E6B8B7', 'bc': '963634', 'b': True, 'align': 'right',
             'bb': 1},
            {'name': 'Заголовок 6', 'bg': 'D8E4BC', 'bc': '76933C', 'b': True, 'align': 'right',
             'bb': 1},
            {'name': 'Заголовок 7', 'bg': 'CCC0DA', 'bc': '60497A', 'b': True, 'align': 'right',
             'bb': 1},
            {'name': 'Заголовок 8', 'bg': 'B7DEE8', 'bc': '31869B', 'b': True, 'align': 'right',
             'bb': 1},
            {'name': 'Итоги 1', 'bg': 'FABF8F', 'bc': 'E26B0A', 'b': True, 'align': 'right',
             'bb': 1, 'bt': 'thick'},
            {'name': 'Таблица 1', 'bl': '1', 'br': '1', 'bt': '1', 'bb': 1, 'wr': True, 'bc': '808080'},
            {'name': 'Шапка 1', 'bl': '1', 'br': '1', 'bt': '1', 'bb': 1, 'b': True}
        ]

        for s in styles:
            n = s['name']
            self.styles[n] = self.wb.add_format()
            self.styles[n].set_font_name(self.font)
            self.styles[n].set_font_size(self.font_size)
            if 'txtclr' in s.keys():
                self.styles[n].set_font_color(s['txtclr'])
            if 'b' in s.keys():
                self.styles[n].set_bold(s['b'])
            if 'itl' in s.keys():
                self.styles[n].set_italic(s['itl'])
            if 'align' in s.keys():
                self.styles[n].set_align(s['align'])
            if 'valign' in s.keys():
                self.styles[n].set_align(s['valign'])
            if 'wr' in s.keys():
                self.styles[n].set_text_wrap(s['wr'])
            if 'bg' in s.keys():
                self.styles[n].set_pattern(1)
                self.styles[n].set_bg_color(s['bg'])
            if 'bb' in s.keys():
                self.styles[n].set_bottom(s['bb'])
                self.styles[n].set_bottom_color(s.get('bc', '000000'))
            if 'bt' in s.keys():
                self.styles[n].set_top(s['bt'])
                self.styles[n].set_top_color(s.get('bc', '000000'))
            if 'br' in s.keys():
                self.styles[n].set_right(s['br'])
                self.styles[n].set_right_color(s.get('bc', '000000'))
            if 'bl' in s.keys():
                self.styles[n].set_left(s['bl'])
                self.styles[n].set_left_color(s.get('bc', '000000'))

    def save(self, pathname='test.xlsx'):
        """Присваивает имя файлу и закрывает"""
        # self.wb.filename('test.xlsx')
        self.wb.close()

    def column_size(self, clm, sz):
        """Устанавливает размеры столбцов"""

        if type(sz) not in (list, tuple):
            col = num_to_col(sz)
            self.ws.column_dimensions[col].width = sz + 0.7109375
        else:
            for i, cs in enumerate(sz):
                col = num_to_col(clm + i)
                self.ws.column_dimensions[col].width = cs + 0.7109375

    def row_size(self, rw, sz):
        """Устанавливает размеры строк"""
        # TODO: адаптировать под новый модуль
        if type(sz) != list:
            self.wb.ActiveSheet.Rows(rw).RowHeight = sz
        else:
            for (val, i) in zip(sz, range(len(sz))):
                self.wb.ActiveSheet.Rows(rw + i).RowHeight = val

    def column_fit(self, clm, ln=1):
        """Установка авторазмера колонок"""
        # TODO: адаптировать под новый модуль
        cells = num_to_col(clm) + ":" + num_to_col(clm + ln - 1)
        self.wb.ActiveSheet.Columns(cells).EntireColumn.AutoFit

    def row_fit(self, rw, ln=1):
        """Установка авторазмера строк"""
        # TODO: адаптировать под новый модуль
        cells = str(rw) + ":" + str(rw + ln - 1)
        self.wb.ActiveSheet.Rows(cells).EntireRow.AutoFit

    def paint_cells(self, range_cells, ink=None, fill=None):
        """Раскрашивание ячеек
        cells - диапазон ячеек формата 'A1:B3'
        ink - цвет шрифта в формате FFF000
        paper - цвет заливки в формате FFF000
        """
        cells = self.ws[range_cells]
        filling = PatternFill(fill_type='solid', start_color=fill, end_color=fill)
        font_color = Font(color=ink)
        if type(cells) is tuple:
            for cell in cells:
                for i in cell:
                    if ink:
                        i.font = font_color
                    if fill:
                        i.fill = filling
        else:
            if ink:
                cells.font = font_color
            if fill:
                cells.fill = filling

    def formatting(self, rw, clm, ha=None, va=None, wrap=None, bld=None, itl=None,
                   nf=None, rot=None, sz=None):
        """Выравнивание текста в ячейках по горизонтали и вертикали; перенос текста; формат числа
        :param int rw: row - Строка
        :param int clm: column - Колонка
        :param str ha: horizontal align = l - xlLeft, r - xlRight, c - xlCenter
        :param str va: vertical align =  t - xlTop, c - xlCenter, b - xlBottom
        :param str wrap: f - False, t - True
        :param list|tuple|str nf: NumberFormat "General" по умолчанию "Общий". Передаётся списком
        :param list|tuple|int rot: rotate: Ориентация текста (поворот в градусах)
        :param str bld: bold жирный текст f-False, t-True
        :param list|tuple|int sz: size font Размер шрифта
        :param str itl: italic Курсивный шрифт
        """
        align = {'l': 'left', 'r': 'right', 'c': 'center', 't': 'top', 'bld': 'bottom'}
        ft = {'f': False, 't': True}
        if ha:
            ha = ha.lower()
        if va:
            va = va.lower()
        if wrap:
            wrap = wrap.lower()
        if bld:
            bld = bld.lower()
        if type(rot) not in (list, tuple):
            rot = (rot,)
        if type(sz) not in [list, tuple]:
            sz = (sz,)
        if type(nf) not in [list, tuple]:
            nf = (nf,)
        lst = [ha, va, wrap, bld, itl, nf, rot, sz]
        rw_len = len(max((i for i in lst if i is not None), key=len))
        for i in range(rw_len):
            cells = num_to_col(clm + i) + str(rw)
            h = align[ha[min(i, len(ha) - 1)]] if ha else None
            v = align[va[min(i, len(va) - 1)]] if va else None
            w = ft[wrap[min(i, len(wrap) - 1)]] if wrap else None
            b = ft[bld[min(i, len(bld) - 1)]] if bld else self.ws[cells].font.b
            it = ft[itl[min(i, len(itl) - 1)]] if itl else self.ws[cells].font.i
            r = rot[min(i, len(rot) - 1)] if any(rot) else self.ws[cells].alignment.textRotation
            fz = sz[min(i, len(sz) - 1)] if any(sz) else self.ws[cells].font.sz
            n = nf[min(i, len(nf) - 1)] if any(nf) else self.ws[cells].number_format

            self.ws[cells].alignment = Alignment(horizontal=h, vertical=v,
                                                 text_rotation=r,
                                                 wrap_text=w)
            self.ws[cells].font = Font(size=fz, bold=b, italic=it)
            self.ws[cells].number_format = n

    def grid_set(self, wt='xlThin', ls='xlContinuous', tc=2):
        """Настройка сетки"""
        self.theme_color = tc
        # Weight - толщины линий
        weight = {'xlThin': 2, 'xlHairline': 1, 'xlThick': 4, 'xlMedium': -4138}
        self.weight = weight[wt]
        # LineStyle - стиль линий
        line_style = {'xlContinuous': 1, 'xlDot': -4118, 'xlDash': -4115, 'xlDashDot': 4, 'xlNone': -4142,
                      'xlDouble': -4119, 'xlSlantDashDot': 13}
        self.line_style = line_style[ls]

    def grid(self, rw=0, clm=0, ln=0, hg=0, sd='lrud'):
        """Отрисовка сетки"""
        # TODO: адаптировать под новый модуль
        cells = num_to_col(clm) + str(rw) + ":" + num_to_col(clm + abs(ln) - 1) + str(rw + abs(hg) - 1)
        sd = sd.lower()
        side = {'l': 7, 'u': 8, 'd': 9, 'r': 10, 'v': 11, 'h': 12}
        for i in sd:
            gr = self.wb.ActiveSheet.Range(cells).Borders(side[i])
            gr.LineStyle = self.line_style
            gr.ThemeColor = self.theme_color
            gr.Weight = self.weight
            gr.TintAndShade = 0

    def move_sheet(self, s1, s2):
        """Перемещает лист s1 перед s2"""
        # TODO: адаптировать под новый модуль
        self.wb.Sheets(s1).Move(Before=self.wb.Sheets(s2))

    def pic_insert(self, rw, clm, ln=1, hg=1, file='', hor='c', ver='c'):
        """Вставка картинки"""
        # TODO: адаптировать под новый модуль
        if not os.path.exists(file):
            return None
        cells = num_to_col(clm) + str(rw) + ":" + num_to_col(clm + abs(ln) - 1) + str(rw + abs(hg) - 1)
        TargetCells = self.wb.ActiveSheet.Range(cells)
        pic = self.wb.ActiveSheet.Shapes.AddPicture(file, 0, 1, -1, -1, -1, -1)
        lWscale = pic.Height / pic.Width
        pic.LockAspectRatio = 0
        pic.Height = TargetCells.Height
        pic.Width = TargetCells.Width
        Aspect = TargetCells.Height / TargetCells.Width
        if lWscale < Aspect:
            pic.ScaleHeight((lWscale / Aspect), 0, 0)
            pic.ScaleHeight(0.9, 0, 0)
            pic.ScaleWidth(0.9, 0, 0)
        else:
            pic.ScaleWidth((Aspect / lWscale), 0, 0)
        hor = hor.lower()
        ver = ver.lower()
        ph = pic.Height
        pw = pic.Width
        tch = TargetCells.Height
        tcw = TargetCells.Width
        hl = TargetCells.Left
        hc = TargetCells.Left + (tcw - pw) / 2
        hr = TargetCells.Left + (tcw - pw)
        vt = TargetCells.Top
        vc = TargetCells.Top + (tch - ph) / 2
        vb = TargetCells.Top + (tch - ph)
        h_align = {'l': hl, 'c': hc, 'r': hr}
        v_align = {'t': vt, 'c': vc, 'b': vb}
        pic.Top = v_align[ver]
        pic.Left = h_align[hor]
        pic.LockAspectRatio = 1

    def write(self, row, col, *args):
        self.ws.write(row, col, *args)

    def put_val(self, row=1, column=1, value=''):
        """Запись данных в ячейки"""

        if type(value) in (list, tuple):
            for col, val in enumerate(value):
                self.ws.cell(row, column + col, val)
        else:
            self.ws.cell(row, column, value)
        return row + 1

    def print_area(self, rw, clm, ln, hg):
        """Задаёт область печати"""
        # TODO: адаптировать под новый модуль
        cells = num_to_col(clm) + str(rw) + ":" + num_to_col(clm + abs(ln) - 1) + str(rw + abs(hg) - 1)
        self.wb.ActiveSheet.PageSetup.PrintArea = cells

    def table_style(self, rw, clm, ln, hg, tabname, style='TableStyleLight2', st=True):
        """Задать таблицу стилей"""
        # TODO: адаптировать под новый модуль
        cells = num_to_col(clm) + str(rw) + ":" + num_to_col(clm + abs(ln) - 1) + str(rw + abs(hg) - 1)
        range = self.wb.ActiveSheet.Range(cells)
        self.wb.ActiveSheet.ListObjects.Add(1, range, 0, 1).Name = tabname
        self.wb.ActiveSheet.ListObjects(tabname).TableStyle = style
        self.wb.ActiveSheet.ListObjects(tabname).ShowTotals = st
        range.AutoFilter()

    def format_cond(self, rang, tp, op, f1, f2):
        """Условное форматирование
           rang - диапозон ("A:C")
           tp - Type Определяет, основан ли условный формат на значении ячейки или выражении
           op - Operator Условный оператор формата. Может быть одна из следующих констант XlFormatConditionOperator:
                xlBetween, xlEqual, xlGreater, xlGreaterEqual, xlLess, xlLessEqual, xlNotBetween, или xlNotEqual.
                Если Тип - xlExpression, аргумент Оператора проигнорирован
           f1 - Formula1 Значение или выражение связанное с условным форматированием. Может быть постоянное значение, строковое значение, ссылка на ячейку или формула.
           f2 - Formula2 Значение или выражение связанное со второй частью условного форматирования, когда Оператор - xlBetween или xlNotBetween (иначе, этот параметр проигнорирован).
                Может быть постоянное значение, строковое значение, ссылка на ячейку или формула.
        """
        # TODO: адаптировать под новый модуль
        self.wb.ActiveSheet.Range(rang).FormatConditions.Add(tp, op, f1, f2)

    def style_to_range(self, rang, style):
        """Применение стиля к диапазону ячеек"""

        cells = self.ws[rang]
        for rng in cells:
            for cell in rng:
                cell.style = style

    def named_ranges(self, name, attr_txt):
        new_range = openpyxl.workbook.defined_name.DefinedName(name, attr_text=attr_txt)
        self.wb.defined_names.append(new_range)
        return

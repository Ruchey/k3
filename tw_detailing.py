# -*- coding: utf-8 -*-

import os
from collections import namedtuple, OrderedDict

import k3r


class Report:
    def __init__(self, db, cnt_scrap=50, scrap_rate=2, margin=3, block_res=True, client=True, joiners=True):
        """
        @param db: База данных
        @type cnt_scrap: Колличество фурнитуры, при котором учитывается коэф. брака
        @param scrap_rate: Коэффициент брака на фурнитуру. Используется при кол-ве более 50шт
        @type margin: Коэффициент наценки
        @param block_res: Блок результатов на главном листе. Вкл или Выкл.
        @param client: Лист для клиентов
        @param joiners: Лист для столяров
        """
        self.cnt_scrap = cnt_scrap
        self.scrap_rate = scrap_rate
        self.margin = margin
        self.block_res = block_res
        self.client = client
        self.joiners = joiners
        self.nm = k3r.nomenclature.Nomenclature(db)
        self.bs = k3r.base.Base(db)
        self.pn = k3r.panel.Panel(db)
        self.ln = k3r.long.Long(db)
        self.pr = k3r.prof.Profile(db)
        self.bs = k3r.base.Base(db)
        self.xl = k3r.xl.Doc()
        self.row = 1
        self.RUB = self.xl.F_RUB
        self.sum_mat_cells = []
        self.links_tab_sheets = []
        self.links_tab_long = []
        self.links_tab_profiles = []
        self.links_tab_acc = []

    def cost(self):
        cost = '=G{0}*F{0}'.format(self.row)
        return cost

    def put_val(self, val):
        self.row = self.xl.put_val(self.row, 1, val)

    def section_heading(self, row, name, style):
        self.row = self.xl.put_val(row, 1, name)
        self.xl.ws.merge_cells('A{0}:H{0}'.format(row))
        self.xl.style_to_range('A{0}:H{0}'.format(row), style)

    def format_head_mat(self, row):
        self.xl.style_to_range('A{0}:H{0}'.format(row), 'Заголовок 6')
        self.xl.formatting(row, 1, ha='lllllc', va='c', bld='f',
                           nf=('@', '@', '@', '@', '@', '0.0', self.RUB, self.RUB))
        self.xl.ws.merge_cells('A{0}:D{0}'.format(row))

    def format_head_mat2(self, row):
        self.xl.style_to_range('A{0}:H{0}'.format(row), 'Шапка 1')
        self.xl.formatting(row, 1, ha='lllllc', va='tc', bld='t',
                           nf=('@', '@', '@', '@', '@', '0.0', self.RUB, self.RUB))
        self.xl.ws.merge_cells('A{0}:D{0}'.format(row))

    def header(self, cs=(4, 47, 5, 5, 6.71, 5.3, 10, 10), pg=1):
        to = self.bs.torderinfo()
        number = to.ordernumber
        name = to.ordername if to.ordername else ''
        customer = to.customer if to.customer else ''
        phone = to.telephonenumber if to.telephonenumber else ''
        adr = to.address if to.address else ''
        executor = to.executor if to.executor else ''
        acceptor = to.acceptor if to.acceptor else ''
        addinfo = to.additionalinfo if to.additionalinfo else ''
        data = to.orderexpirationdata if to.orderexpirationdata.year > 2000 else ''
        val1 = [
            ('Заказ №{0}'.format(number)),
            ('Название: {0}'.format(name)),
            ('Заказчик: {0} Телефон: {1}'.format(customer, phone)),
            ('Адрес: {0}'.format(adr)),
            ('Приёмщик: {0} Исполнитель: {1}'.format(acceptor, executor)),
            ('Дата исполнения: {}'.format(data.strftime("%m.%d.%y")), '', 'Собирать в цехе', '', '', 'корпусном'),
            ('Дополнительно: {}'.format(addinfo))
        ]
        val2 = [
            ('Заказ №{0}'.format(number)),
            ('Приёмщик: {0} Исполнитель: {1}'.format(acceptor, executor)),
            ('Дополнительно: {}'.format(addinfo))
        ]
        val = {1: val1, 2: val1, 3: val2}
        for v in val[pg]:
            self.xl.formatting(self.row, 1, sz=12, bld='f', itl='f')
            self.put_val(v)
        if pg == 1:
            dv = self.xl.dv(type="list", formula1='"корпусном,металлическом,мягкой мебели,столярном"', allow_blank=True)
            dv.error = 'Запись отсутствует в списке'
            dv.errorTitle = 'Ошибочный ввод'
            dv.prompt = 'Выберите цех'
            dv.promptTitle = 'Список цехов'
            dv.add(self.xl.ws['F{}'.format(self.row - 2)])
            self.xl.ws.add_data_validation(dv)
            self.xl.ws.merge_cells('C{0}:E{0}'.format(self.row - 2))
            self.xl.ws.merge_cells('F{0}:G{0}'.format(self.row - 2))
            self.xl.ws.merge_cells('A{0}:H{0}'.format(self.row - 1))
        self.row += 1
        self.xl.col_size(1, cs)

    def get_pans(self, pans):
        keys = ('cpos', 'name', 'length', 'width', 'cnt')
        l_pans = []
        for id_pan in pans:
            Pans = namedtuple('Pans', keys)
            telems = self.bs.telems(id_pan)
            cpos = telems.commonpos
            name = telems.name
            name += ' (Г)' if self.pn.form(id_pan) > 0 else ''
            name += ' (К)' if self.pn.curvepath(id_pan) > 0 else ''
            length = self.pn.length(id_pan)
            width = self.pn.width(id_pan)
            cnt = telems.count
            pans = Pans(cpos, name, length, width, cnt)
            l_pans.append(pans)
        new_list = k3r.utils.group_by_key(l_pans, 'cpos', 'cnt')
        return new_list

    def det_tab_sheets(self, tpp=None):
        mat_id = self.nm.mat_by_uid(2, tpp)
        if not mat_id:
            return
        self.section_heading(self.row, 'Листовые материалы', 'Заголовок 2')
        val = ('№', 'Наименование', 'Д-на', 'Ш-на', 'шт', 'Всего', 'Цена', 'Стоимость')
        self.put_val(val)
        self.xl.style_to_range('A{0}:H{0}'.format(self.row - 1), 'Заголовок 3')
        self.xl.formatting(self.row - 1, 1, ha='cccccccc', va='c')
        mats = []
        for i in mat_id:
            prop = self.nm.properties(i)
            mats.append(prop)
        mats.sort(key=lambda x: x.thickness, reverse=True)
        for mat in mats:
            pans = self.get_pans(self.pn.list_panels(mat.id, tpp))
            pic = getattr(mat, 'pictures', '')
            wc = mat.wastecoeff
            wood = getattr(mat, 'wood', 1)
            total = '=SUM(F{0}:F{1})*{2}'.format(self.row + 1, self.row + len(pans), wc)
            val = (mat.name, '', '', '', mat.unitsname, total, mat.price, self.cost())
            self.sum_mat_cells.append('H{}'.format(self.row))
            self.put_val(val)
            self.format_head_mat(self.row - 1)
            start_row = self.row
            for p in pans:
                area = '=ROUND(C{0}*D{0}*E{0}/10^6,2)'.format(self.row)
                val = (p.cpos, p.name, p.length, p.width, p.cnt, area)
                self.put_val(val)
            self.xl.style_to_range('A{0}:H{1}'.format(start_row, self.row - 1), 'Таблица 1')
            self.xl.paint_cells('F{0}:F{1}'.format(start_row, self.row - 1), ink='B5B5B5')
            self.xl.ws.row_dimensions.group(start_row, self.row - 1, hidden=True)
            self.links_tab_sheets.append({'head': start_row - 1, 'det': (start_row, self.row),
                                          'pic': pic, 'wd': wood})

    def det_tab_long(self):
        long_list = self.ln.long_list()
        if not long_list:
            return
        self.section_heading(self.row, 'Длиномеры', 'Заголовок 2')
        val = ('Наименование', '', 'Д-на', 'Ш-на', 'шт', 'Всего', 'Цена', 'Стоимость')
        self.put_val(val)
        self.xl.style_to_range('A{0}:H{0}'.format(self.row - 1), 'Заголовок 3')
        self.xl.ws.merge_cells('A{0}:B{0}'.format(self.row - 1))
        self.xl.formatting(self.row - 1, 1, ha='cccccccc', va='c')
        total = self.ln.total()
        for i in total:
            filter_long = list(lg for lg in long_list if lg.type == i.type and lg.matid == i.matid
                               and lg.goodsid == i.goodsid)
            prop = self.nm.properties(i.matid)
            goods = self.bs.tngoods(i.goodsid)
            wood = getattr(prop, 'wood', 1)
            pic = getattr(prop, 'pictures', '')
            name = '{0} {1} {2}'.format(goods.groupname, goods.name, prop.name)
            name = ' '.join(OrderedDict((w, 0) for w in name.split()).keys())
            wc = prop.wastecoeff
            # len_wc = '={0}*{1}'.format(i.length, wc)
            total = '=SUM(F{0}:F{1})*{2}'.format(self.row + 1, self.row + len(filter_long), wc)
            val = (name, '', '', '', prop.unitsname, total, prop.price, self.cost())
            self.sum_mat_cells.append('H{}'.format(self.row))
            self.put_val(val)
            self.format_head_mat(self.row - 1)
            start_row = self.row
            for j in filter_long:
                prop = self.nm.properties(j.matid)
                name = prop.name
                width = j.width
                if j.form > 0:
                    name += ' (R) гл. {0}мм'.format(width)
                if j.type in (2, 5, 6, 7):
                    width = j.height
                sum_len = '=ROUND(C{0}*E{0}/1000,2)'.format(self.row)
                val = (name, '', j.length, width, j.cnt, sum_len)
                self.put_val(val)
                self.xl.ws.merge_cells('A{0}:B{0}'.format(self.row - 1))
            self.xl.ws.row_dimensions.group(start_row, self.row - 1, hidden=False)
            self.xl.style_to_range('A{0}:H{1}'.format(start_row, self.row - 1), 'Таблица 1')
            self.xl.paint_cells('F{0}:F{1}'.format(start_row, self.row - 1), ink='B5B5B5')
            self.links_tab_long.append({'head': start_row - 1, 'det': (start_row, self.row),
                                        'pic': pic, 'wd': wood})

    def det_tab_profiles(self, tpp=None):
        profiles_list = self.pr.profiles(tpp=tpp)
        if not profiles_list:
            return
        self.section_heading(self.row, 'Профиля', 'Заголовок 2')
        val = ('Наименование', '', '', '', 'арт', 'Всего', 'Цена', 'Стоимость')
        self.put_val(val)
        self.xl.style_to_range('A{0}:H{0}'.format(self.row - 1), 'Заголовок 3')
        self.xl.ws.merge_cells('A{0}:D{0}'.format(self.row - 1))
        self.xl.formatting(self.row - 1, 1, ha='cccccccc', va='c')
        for i in self.pr.total():
            filter_prof = list(pr for pr in profiles_list if pr.priceid == i.priceid)
            prop = self.nm.properties(i.priceid)
            wood = getattr(prop, 'wood', 1)
            upl_prof = getattr(prop, 'upl_prof', 0)
            pic = getattr(prop, 'pictures', '')
            wc = prop.wastecoeff
            len_wc = '={0}*{1}'.format(i.length, wc)
            val = (prop.name, '', '', '', prop.unitsname, len_wc, prop.price, self.cost())
            self.sum_mat_cells.append('H{}'.format(self.row))
            self.put_val(val)
            self.format_head_mat(self.row - 1)
            start_row = self.row
            end_row = self.row + len(filter_prof) - 1
            self.xl.style_to_range('A{0}:H{1}'.format(self.row, end_row), 'Таблица 1')
            if not upl_prof:
                for j in filter_prof:
                    name = prop.name
                    if j.formtype > 0:
                        name += ' (R)'
                    val = (name, '', '', '', prop.article, '', j.length, j.cnt)
                    self.put_val(val)
                    self.xl.ws.merge_cells('A{0}:D{0}'.format(self.row - 1))
                    self.xl.formatting(self.row - 1, 1, ha='llllllrc', va='c',
                                       nf=('@', '@', '@', '@', '@', '@', '0\\m\\m', '#шт'))
            self.xl.ws.row_dimensions.group(start_row, self.row - 1, hidden=True)
            self.links_tab_profiles.append({'head': start_row - 1, 'det': (start_row, self.row),
                                            'pic': pic, 'wd': wood})

    def det_tab_acc(self, tpp=None):
        acc = self.nm.acc_by_uid(tpp=tpp)
        acc_long = self.nm.acc_long(tpp=tpp)
        if not acc and not acc_long:
            return
        self.section_heading(self.row, 'Комплектующие', 'Заголовок 2')
        start_row = self.row
        end_row = self.row + len(acc) + len(acc_long) - 1
        self.xl.style_to_range('A{0}:H{1}'.format(self.row, end_row), 'Таблица 1')
        for ac in acc:
            prop = self.nm.properties(ac.priceid)
            wood = getattr(prop, 'wood', 1)
            pic = getattr(prop, 'pictures', '')
            name = ac.name
            if ac.article:
                name += ' арт.{}'.format(ac.article)
            units_name = ac.unitsname
            ac_cnt = ac.cnt
            if ac.unitsid == 4 and self.scrap_rate > 0 and ac_cnt > self.cnt_scrap:
                units_name += '**'
                ac_cnt *= 1 + self.scrap_rate / 100
            val = (name, '', '', '', units_name, ac_cnt, ac.price, self.cost())
            self.put_val(val)
            self.xl.ws.merge_cells('A{0}:D{0}'.format(self.row - 1))
            self.xl.formatting(self.row - 1, 1, ha='llllccrr', va='c',
                               nf=('@', '@', '@', '@', '@', '0', self.RUB, self.RUB))
            self.links_tab_acc.append({'head': self.row - 1, 'pic': pic, 'wd': wood})
        for ac in acc_long:
            prop = self.nm.properties(ac.priceid)
            wood = getattr(prop, 'wood', 1)
            pic = getattr(prop, 'pictures', '')
            name = ac.name
            if ac.article:
                name += ' арт.{}'.format(ac.article)
            cost = '=D{0}*G{0}*F{0}'.format(self.row)
            val = (name, '', '', ac.length, ac.unitsname, ac.cnt, ac.price, cost)
            self.put_val(val)
            self.xl.ws.merge_cells('A{0}:D{0}'.format(self.row - 1))
            self.xl.formatting(self.row - 1, 1, ha='llllccrr', va='c',
                               nf=('@', '@', '@', '@', '@', '0', self.RUB, self.RUB))
            if wood:
                self.links_tab_acc.append({'head': self.row - 1, 'pic': pic, 'wd': wood})
        self.sum_mat_cells.append('H{0}:H{1}'.format(start_row, end_row))

    def det_tab_bands(self):
        bands = self.nm.bands()
        if not bands:
            return
        self.section_heading(self.row, 'Кромка', 'Заголовок 2')
        val = ('Кромка', '', 'Т-на', 'Д-на', 'к.о.', 'Д*ко', 'Цена', 'Стоимость')
        self.put_val(val)
        self.xl.style_to_range('A{0}:H{0}'.format(self.row - 1), 'Заголовок 3')
        self.xl.formatting(self.row - 1, 1, ha='llc', va='c')
        start_row = self.row
        end_row = self.row + len(bands) - 1
        self.xl.style_to_range('A{0}:H{1}'.format(self.row, end_row), 'Таблица 1')
        for b in bands:
            prop = self.nm.properties(b.id)
            wc = prop.wastecoeff
            len_wc = '=D{0}*E{0}'.format(self.row)
            val = (prop.name, '', b.thick, b.length, wc, len_wc, prop.price, self.cost())
            self.put_val(val)
            self.xl.ws.merge_cells('A{0}:B{0}'.format(self.row - 1))
            self.xl.formatting(self.row - 1, 1, ha='llllrrrr', va='c', bld='f',
                               nf=('@', '@', '0', '0.0', '0.0', '0.0', self.RUB, self.RUB))
        self.sum_mat_cells.append('H{0}:H{1}'.format(start_row, end_row))

    def det_results(self):
        res_sum = '=SUM({0})'.format(','.join(self.sum_mat_cells))
        val = ('Материалы по заказу:', '', '', res_sum)
        res_sum_name = '{0}!$H${1}'.format(self.xl.ws.title, self.row)
        self.xl.named_ranges('res_sum', res_sum_name)
        self.row = self.xl.put_val(self.row, 5, val)
        self.xl.style_to_range('A{0}:H{0}'.format(self.row - 1), 'Итоги 1')
        self.xl.formatting(self.row - 1, 5, ha='lllr', va='c', bld='tf', nf=('@', '@', '@', self.RUB))
        if not self.block_res:
            return
        if self.scrap_rate > 0:
            val = ('** - посчитано с запасом {}%'.format(self.scrap_rate))
            self.row = self.xl.put_val(self.row, 2, val)
        self.margin = 3
        res_coeff = '=res_sum*G{0}'.format(self.row)
        marg_sum_name = '{0}!$H${1}'.format(self.xl.ws.title, self.row)
        self.xl.named_ranges('marg_sum', marg_sum_name)
        val = (self.margin, res_coeff)
        self.row = self.xl.put_val(self.row, 7, val)
        self.xl.style_to_range('G{0}:H{0}'.format(self.row - 1), 'Таблица 1')
        self.xl.formatting(self.row - 1, 7, ha='cr', va='c', bld='ft', nf=('x_0.0', self.RUB))
        self.row += 1
        vals = [
            ('Листовой камень', '', '', '', '', 1.5),
            ('Металл', '', '', '', '', 4),
            ('Элементы мягкой мебели', '', '', '', '', 1.5),
            ('Заказные позиции', '', '', '', '', 1.3),
        ]
        if self.joiners:
            vals.insert(0, ('Столярные элементы', '=sum_joiners', '', '', '', 2.5))
        start_row = self.row
        end_row = self.row + len(vals) - 1
        self.xl.style_to_range('B{0}:H{1}'.format(self.row, end_row), 'Таблица 1')
        for val in vals:
            val += ('=C{0}*G{0}'.format(self.row),)
            self.row = self.xl.put_val(self.row, 2, val)
            self.xl.ws.merge_cells('C{0}:F{0}'.format(self.row - 1))
            self.xl.formatting(self.row - 1, 2, ha='lrrrrcr', va='c', bld='f',
                               nf=('@', self.RUB, '@', '@', '@', 'x_0.0', self.RUB))
        self.row += 1
        val = ('Закупочная стоимость', '=SUM(res_sum, C{0}:C{1})'.format(start_row, end_row))
        self.xl.put_val(self.row, 2, val)
        self.xl.ws.merge_cells('C{0}:D{0}'.format(self.row))
        self.xl.style_to_range('B{0}:D{0}'.format(self.row), 'Таблица 1')
        self.xl.formatting(self.row, 2, ha='lrr', va='c', bld='f', nf=('@', self.RUB))
        val = ('Итого:', '=SUM(marg_sum, H{0}:H{1})'.format(start_row, end_row))
        self.xl.put_val(self.row, 7, val)
        self.xl.style_to_range('G{0}:H{0}'.format(self.row), 'Итоги 1')
        self.xl.formatting(self.row, 7, ha='lr', va='c', bld='t', nf=('@', self.RUB))

    def page_joiners(self):
        """Создаёт лист для столяров"""

        self.xl.new_sheet('Столярка')
        self.sum_mat_cells = []
        self.row = 1
        pg = 'Деталировка'
        self.header(pg=2)
        if self.links_tab_sheets:
            self.section_heading(self.row, 'Листовые материалы', 'Заголовок 2')
            val = ('№', 'Наименование', 'Д-на', 'Ш-на', 'шт', 'Всего', 'Цена', 'Стоимость')
            self.put_val(val)
            self.xl.style_to_range('A{0}:H{0}'.format(self.row - 1), 'Заголовок 3')
            self.xl.formatting(self.row - 1, 1, ha='cccccccc', va='c')
            for i in self.links_tab_sheets:
                row = i['head']
                val = (
                    '={0}!A{1}'.format(pg, row), '', '', '', '={0}!E{1}'.format(pg, row), '={0}!F{1}'.format(pg, row),
                    '={0}!G{1}'.format(pg, row), '={0}!H{1}'.format(pg, row))
                self.sum_mat_cells.append('H{}'.format(self.row))
                self.put_val(val)
                self.format_head_mat(self.row - 1)
                start_row = self.row
                for j in range(i['det'][0], i['det'][1]):
                    row = j
                    val = ('={0}!A{1}'.format(pg, row), '={0}!B{1}'.format(pg, row), '={0}!C{1}'.format(pg, row),
                           '={0}!D{1}'.format(pg, row), '={0}!E{1}'.format(pg, row), '={0}!F{1}'.format(pg, row))
                    self.put_val(val)
                self.xl.style_to_range('A{0}:H{1}'.format(start_row, self.row - 1), 'Таблица 1')
                self.xl.paint_cells('F{0}:F{1}'.format(start_row, self.row - 1), ink='B5B5B5')
                self.xl.ws.row_dimensions.group(start_row, self.row - 1, hidden=True)

        if self.links_tab_long:
            self.section_heading(self.row, 'Длиномеры', 'Заголовок 2')
            val = ('Наименование', '', 'Д-на', 'Ш-на', 'шт', 'Всего', 'Цена', 'Стоимость')
            self.put_val(val)
            self.xl.style_to_range('A{0}:H{0}'.format(self.row - 1), 'Заголовок 3')
            self.xl.ws.merge_cells('A{0}:B{0}'.format(self.row - 1))
            self.xl.formatting(self.row - 1, 1, ha='cccccccc', va='c')
            for i in self.links_tab_long:
                row = i['head']
                val = (
                    '={0}!A{1}'.format(pg, row), '', '', '', '={0}!E{1}'.format(pg, row), '={0}!F{1}'.format(pg, row),
                    '={0}!G{1}'.format(pg, row), '={0}!H{1}'.format(pg, row))
                self.sum_mat_cells.append('H{}'.format(self.row))
                self.put_val(val)
                self.format_head_mat2(self.row - 1)
                if i['pic']:
                    self.xl.row_size(self.row - 1, 150)
                    self.xl.pic_insert(self.row - 1, 2, path=i['pic'], max_col=3, align='r', valign='b', cf=0.8, px=6)
                start_row = self.row
                for j in range(i['det'][0], i['det'][1]):
                    row = j
                    val = (
                        '={0}!A{1}'.format(pg, row), '', '={0}!C{1}'.format(pg, row),
                        '={0}!D{1}'.format(pg, row), '={0}!E{1}'.format(pg, row),
                        '={0}!F{1}'.format(pg, row))
                    self.put_val(val)
                    self.xl.ws.merge_cells('A{0}:B{0}'.format(self.row - 1))
                if i['pic']:
                    self.xl.row_size(self.row, 9)
                    self.row += 1
                self.xl.style_to_range('A{0}:H{1}'.format(start_row, self.row - 2), 'Таблица 1')
                self.xl.paint_cells('F{0}:F{1}'.format(start_row, self.row - 2), ink='B5B5B5')
                self.xl.ws.row_dimensions.group(start_row, self.row - 2, hidden=False)

        if self.links_tab_profiles:
            self.section_heading(self.row, 'Профиля', 'Заголовок 2')
            val = ('Наименование', '', '', '', 'арт', 'Всего', 'Цена', 'Стоимость')
            self.put_val(val)
            self.xl.style_to_range('A{0}:H{0}'.format(self.row - 1), 'Заголовок 3')
            self.xl.ws.merge_cells('A{0}:D{0}'.format(self.row - 1))
            self.xl.formatting(self.row - 1, 1, ha='cccccccc', va='c')
            for i in self.links_tab_profiles:
                row = i['head']
                val = (
                    '={0}!A{1}'.format(pg, row), '', '', '', '={0}!E{1}'.format(pg, row), '={0}!F{1}'.format(pg, row),
                    '={0}!G{1}'.format(pg, row), '={0}!H{1}'.format(pg, row))
                self.sum_mat_cells.append('H{}'.format(self.row))
                self.put_val(val)
                self.format_head_mat2(self.row - 1)
                if i['pic']:
                    self.xl.row_size(self.row - 1, 150)
                    self.xl.pic_insert(self.row - 1, 2, path=i['pic'], max_col=3, align='r', valign='b', cf=0.8, px=6)
                start_row = self.row
                end_row = self.row + (i['det'][1] - i['det'][0]) - 1
                self.xl.style_to_range('A{0}:H{1}'.format(self.row, end_row), 'Таблица 1')
                for j in range(i['det'][0], i['det'][1]):
                    row = j
                    val = ('={0}!A{1}'.format(pg, row), '', '', '', '={0}!E{1}'.format(pg, row), '',
                           '={0}!G{1}'.format(pg, row), '={0}!H{1}'.format(pg, row))
                    self.put_val(val)
                    self.xl.ws.merge_cells('A{0}:D{0}'.format(self.row - 1))
                    self.xl.formatting(self.row - 1, 1, ha='llllllrc', va='c',
                                       nf=('@', '@', '@', '@', '@', '@', '0\\m\\m', '#шт'))
                if i['pic']:
                    self.xl.row_size(self.row, 9)
                    self.row += 1
                self.xl.paint_cells('F{0}:F{1}'.format(start_row, self.row - 1), ink='B5B5B5')
                self.xl.ws.row_dimensions.group(start_row, self.row - 1, hidden=False)

        if self.links_tab_acc:
            self.section_heading(self.row, 'Комплектующие', 'Заголовок 2')
            start_row = self.row
            end_row = self.row + len(self.links_tab_acc) - 1
            self.xl.style_to_range('A{0}:H{1}'.format(self.row, end_row), 'Таблица 1')
            for i in self.links_tab_acc:
                row = i['head']
                val = (
                    '={0}!A{1}'.format(pg, row), '', '', '', '={0}!E{1}'.format(pg, row), '={0}!F{1}'.format(pg, row),
                    '={0}!G{1}'.format(pg, row), '={0}!H{1}'.format(pg, row))
                self.put_val(val)
                self.xl.ws.merge_cells('A{0}:D{0}'.format(self.row - 1))
                self.xl.formatting(self.row - 1, 1, ha='llllccrr', va='tc',
                                   nf=('@', '@', '@', '@', '@', '0', self.RUB, self.RUB))
                if i['pic']:
                    self.xl.row_size(self.row - 1, 150)
                    self.xl.pic_insert(self.row - 1, 2, path=i['pic'], max_col=3, align='r', valign='b', cf=0.8, px=6)
            self.sum_mat_cells.append('H{0}:H{1}'.format(start_row, end_row))

        if self.sum_mat_cells:
            self.row += 1
            res_sum = '=SUM({0})'.format(','.join(self.sum_mat_cells))
            val = ('Материалы по заказу:', '', '', res_sum)
            res_sum_name = '{0}!$H${1}'.format(self.xl.ws.title, self.row)
            self.xl.named_ranges('sum_joiners', res_sum_name)
            self.row = self.xl.put_val(self.row, 5, val)
            self.xl.style_to_range('A{0}:H{0}'.format(self.row - 1), 'Итоги 1')
            self.xl.formatting(self.row - 1, 5, ha='lllr', va='c', bld='tf', nf=('@', '@', '@', self.RUB))

    def page_client(self):
        """Создаёт лист для клиента"""

        list_mats = self.links_tab_sheets
        list_mats.extend(self.links_tab_long)
        list_mats.extend(self.links_tab_profiles)
        list_mats.extend(self.links_tab_acc)

        self.xl.new_sheet('Клиент')
        self.row = 1
        self.header(cs=(47, 47), pg=3)
        pg = 'Деталировка'
        i = 1
        if self.links_tab_sheets:
            for vl in self.links_tab_sheets:
                pic = vl['pic']
                if pic:
                    row = vl['head']
                    val = ('={0}!A{1}'.format(pg, row))
                    if i % 2 != 0:
                        self.xl.put_val(self.row, 1, val)
                        self.xl.row_size(self.row, 150)
                        self.xl.pic_insert(self.row, 1, path=pic, align='r', valign='b', cf=0.7, px=6)
                        self.xl.style_to_range('A{0}:B{0}'.format(self.row), 'Таблица 1')
                        self.xl.formatting(self.row, 1, ha='ll', va='tt', wrap='tt')
                    else:
                        self.xl.put_val(self.row, 2, val)
                        self.xl.pic_insert(self.row, 2, path=pic, align='r', valign='b', cf=0.7, px=6)
                        self.row += 1
                i += 1

    def detailing(self):
        """Создаёт основной лист деталировки"""
        self.xl.new_sheet('Деталировка')
        self.header()
        self.det_tab_sheets()
        self.det_tab_long()
        self.det_tab_profiles()
        self.det_tab_acc()
        self.det_tab_bands()
        self.det_results()

    def create(self):
        """Стартовая функция создающая нужные листы"""
        self.detailing()
        if self.joiners:
            self.page_joiners()
        if self.client:
            self.page_client()


if __name__ == '__main__':
    num = 1
    fileDB = r'd:\К3\Самара\Самара черновик\{0}\{0}.mdb'.format(num)
    proj_rep_path = r'd:\К3\Самара\Самара черновик\{0}\Reports'.format(num)
    project = "Деталировка"
    file_path = os.path.join(proj_rep_path, '{}.xlsx'.format(project))
    base = k3r.db.DB()
    base.open(fileDB)
    rep = Report(base)
    rep.create()
    rep.xl.save(file_path)
    base.close()
    os.startfile(file_path)

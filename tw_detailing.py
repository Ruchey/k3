# -*- coding: utf-8 -*-

import os
from collections import namedtuple, OrderedDict

import k3r
from pprint import pprint


class Report:

    def __init__(self, db, scrap_rate=0, block_res=True, client=True, joiners=True):
        """
        @param db: База данных
        @param scrap_rate: Коэффициент брака на фурнитуру. Используется при кол-ве более 50шт
        @param block_res: Блок результатов на главном листе. Вкл или Выкл.
        @param client: Лист для клиентов
        @param joiners: Лист для столяров
        """
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
        self.scrap_rate = scrap_rate
        self.block_res = block_res
        self.client = client
        self.joiners = joiners

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

    def prep(self):
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
        val = [
            ('Заказ №{0}'.format(number)),
            ('Название: {0}'.format(name)),
            ('Заказчик: {0}'.format(customer)),
            ('Телефон: {0}'.format(phone)),
            ('Адрес: {0}'.format(adr)),
            ('Приёмщик: {0} Исполнитель: {1}'.format(acceptor, executor)),
            ('Дата исполнения: {}'.format(data)),
            ('Дополнительно: {}'.format(addinfo))
        ]
        for v in val:
            self.xl.formatting(self.row, 1, sz=[12], bld='tfftf', itl='ftfft')
            self.put_val(v)
        self.row += 1
        cs = [4, 47, 5, 5, 6.71, 5.3, 10, 10]
        self.xl.column_size(1, cs)

    def tab_sheets(self, tpp=None):
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
            wc = prop.wastecoeff
            total = '=SUM(F{0}:F{1})*{2}'.format(self.row+1, self.row+len(pans), wc)
            val = (mat.name, '', '', '', mat.unitsname, total, mat.price, self.cost())
            self.sum_mat_cells.append('H{}'.format(self.row))
            self.put_val(val)
            self.format_head_mat(self.row-1)
            start_row = self.row
            for p in pans:
                area = '=ROUND(C{0}*D{0}*E{0}/10^6,2)'.format(self.row)
                val = (p.cpos, p.name, p.length, p.width, p.cnt, area)
                self.put_val(val)
            self.xl.style_to_range('A{0}:H{1}'.format(start_row, self.row-1), 'Таблица 1')
            self.xl.paint_cells('F{0}:F{1}'.format(start_row, self.row-1), ink='B5B5B5')
            self.xl.ws.row_dimensions.group(start_row, self.row-1, hidden=True)

    def tab_long(self):
        long_list = self.ln.long_list()
        if not long_list:
            return
        self.section_heading(self.row, 'Длиномеры', 'Заголовок 2')
        val = ('№', 'Наименование', 'Д-на', 'Ш-на', 'шт', 'Всего', 'Цена', 'Стоимость')
        self.put_val(val)
        self.xl.style_to_range('A{0}:H{0}'.format(self.row - 1), 'Заголовок 3')
        self.xl.formatting(self.row - 1, 1, ha='cccccccc', va='c')
        total = self.ln.total()
        for i in total:
            filter_long = list(lg for lg in long_list if lg.type == i.type and lg.matid == i.matid
                               and lg.goodsid == i.goodsid)
            prop = self.nm.properties(i.matid)
            goods = self.bs.tngoods(i.goodsid)
            name = '{0} {1} {2}'.format(goods.groupname, goods.name, prop.name)
            name = ' '.join(OrderedDict((w, 0) for w in name.split()).keys())
            wc = prop.wastecoeff
            len_wc = '={0}*{1}'.format(i.length, wc)
            val = (name, '', '', '', prop.unitsname, len_wc, prop.price, self.cost())
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
                val = (name, '', j.length, width, j.cnt)
                self.put_val(val)
                self.xl.ws.merge_cells('A{0}:B{0}'.format(self.row - 1))
            self.xl.ws.row_dimensions.group(start_row, self.row - 1, hidden=False)
            self.xl.style_to_range('A{0}:H{1}'.format(start_row, self.row - 1), 'Таблица 1')

    def tab_profiles(self):
        profiles_list = self.pr.profiles()
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
            wc = prop.wastecoeff
            len_wc = '={0}*{1}'.format(i.length, wc)
            val = (prop.name, '', '', '', prop.unitsname, len_wc, prop.price, self.cost())
            self.sum_mat_cells.append('H{}'.format(self.row))
            self.put_val(val)
            self.format_head_mat(self.row - 1)
            start_row = self.row
            end_row = self.row + len(filter_prof) - 1
            self.xl.style_to_range('A{0}:H{1}'.format(self.row, end_row), 'Таблица 1')
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

    def tab_acc(self):
        acc = self.nm.acc_by_uid()
        acc_long = self.nm.acc_long()
        if not acc and not acc_long:
            return
        self.section_heading(self.row, 'Комплектующие', 'Заголовок 2')
        start_row = self.row
        end_row = self.row + len(acc) + len(acc_long) - 1
        self.xl.style_to_range('A{0}:H{1}'.format(self.row, end_row), 'Таблица 1')
        for ac in acc:
            name = ac.name
            if ac.article:
                name += ' арт.{}'.format(ac.article)
            val = (name, '', '', '', ac.unitsname, ac.cnt, ac.price, self.cost())
            self.put_val(val)
            self.xl.ws.merge_cells('A{0}:D{0}'.format(self.row - 1))
            self.xl.formatting(self.row - 1, 1, ha='llllccrr', va='c',
                               nf=('@', '@', '@', '@', '@', '0', self.RUB, self.RUB))
        for ac in acc_long:
            name = ac.name
            if ac.article:
                name += ' арт.{}'.format(ac.article)
            cost = '=D{0}*G{0}*F{0}'.format(self.row)
            val = (name, '', '', ac.length, ac.unitsname, ac.cnt, ac.price, cost)
            self.put_val(val)
            self.xl.ws.merge_cells('A{0}:D{0}'.format(self.row - 1))
            self.xl.formatting(self.row - 1, 1, ha='llllccrr', va='c',
                               nf=('@', '@', '@', '@', '@', '0', self.RUB, self.RUB))
        self.sum_mat_cells.append('H{0}:H{1}'.format(start_row, end_row))

    def tab_bands(self):
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
            self.xl.ws.merge_cells('A{0}:B{0}'.format(self.row-1))
            self.xl.formatting(self.row - 1, 1, ha='llllrrrr', va='c', bld='f',
                               nf=('@', '@', '0', '0.0', '0.0', '0.0', self.RUB, self.RUB))
        self.sum_mat_cells.append('H{0}:H{1}'.format(start_row, end_row))

    def results(self):
        res_sum = '=SUM({0})'.format(','.join(self.sum_mat_cells))
        val = ('Материалы по заказу:', '', '', res_sum)
        res_sum_name = '{0}!$H${1}'.format(self.xl.ws.title, self.row)
        self.xl.named_ranges('res_sum', res_sum_name)
        self.row = self.xl.put_val(self.row, 5, val)
        self.xl.style_to_range('A{0}:H{0}'.format(self.row-1), 'Итоги 1')
        self.xl.formatting(self.row - 1, 5, ha='lllr', va='c', bld='tf', nf=('@', '@', '@', self.RUB))
        val = ('** - посчитано с запасом {}%'.format(self.scrap_rate))
        self.row = self.xl.put_val(self.row, 2, val)
        coeff = 3
        res_coeff = '=res_sum*G{0}'.format(self.row)
        val = (coeff, res_coeff)
        self.row = self.xl.put_val(self.row, 7, val)
        self.xl.formatting(self.row - 1, 7, ha='cr', va='c', bld='ft', nf=('x_0.0', self.RUB))


def start(base, path):
    rep = Report(base)
    rep.xl.new_sheet('Деталировка')
    rep.prep()
    rep.tab_sheets()
    rep.tab_long()
    rep.tab_profiles()
    rep.tab_acc()
    rep.tab_bands()
    rep.results()
    rep.xl.save(path)
    os.startfile(path)


if __name__ == '__main__':
    num = 1
    fileDB = r'd:\К3\Самара\Самара черновик\{0}\{0}.mdb'.format(num)
    proj_rep_path = r'd:\К3\Самара\Самара черновик\{0}\Reports'.format(num)
    project = "Деталировка"
    file_path = os.path.join(proj_rep_path, '{}.xlsx'.format(project))
    fileDB = r'd:\0\71.mdb'
    db = k3r.db.DB()
    db.open(fileDB)
    start(db, file_path)
    db.close()

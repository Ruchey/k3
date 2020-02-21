# -*- coding: utf-8 -*-

import os
from collections import namedtuple, OrderedDict

import k3r
from pprint import pprint


class Report:

    def __init__(self, db):
        self.nm = k3r.nomenclature.Nomenclature(db)
        self.bs = k3r.base.Base(db)
        self.pn = k3r.panel.Panel(db)
        self.ln = k3r.long.Long(db)
        self.bs = k3r.base.Base(db)
        self.xl = k3r.doc.Doc()
        self.row = 1
        self.RUB = self.xl.F_RUB

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
        self.xl.style_to_range('A{0}:H{0}'.format(row), 'Заголовок 1')
        self.xl.txt_format(row, 1, h_align='lllllc', v_align='c', bold='tf',
                           nf=('@', '@', '@', '@', '@', '0.00', self.RUB, self.RUB))
        self.xl.ws.merge_cells('A{0}:E{0}'.format(row))

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
            self.xl.txt_format(self.row, 1, fsz=[12], bold='tfftf', italic='ftfft')
            self.put_val(v)
        self.row += 1
        cs = [4, 48, 5, 5, 6.71, 5.3, 9, 10]
        self.xl.column_size(1, cs)
        val = ('№', 'Наименование', 'Д-на', 'Ш-на', 'шт', 'Всего', 'Цена', 'Стоимость')
        self.put_val(val)
        self.xl.style_to_range('A{0}:H{0}'.format(self.row - 1), 'Заголовок 3')
        self.xl.txt_format(self.row - 1, 1, h_align='cccccccc', v_align='c')

    def tab_sheets(self, tpp=None):
        mat_id = self.nm.mat_by_uid(2, tpp)
        if not mat_id:
            return
        self.section_heading(self.row, 'Листовые материалы', 'Заголовок 2')
        mats = []
        for i in mat_id:
            prop = self.nm.properties(i)
            mats.append(prop)
        mats.sort(key=lambda x: x.thickness, reverse=True)
        for mat in mats:
            pans = self.get_pans(self.pn.list_panels(mat.id, tpp))
            total = '=SUM(F{0}:F{1})'.format(self.row + 1, self.row + len(pans))
            val = (mat.name, '', '', '', '', total, mat.price, self.cost())
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

    def tab_long(self):
        long_list = self.ln.long_list()
        if not long_list:
            return
        self.section_heading(self.row, 'Длиномеры', 'Заголовок 2')
        total = self.ln.total()
        for i in total:
            filter_long = list(lg for lg in long_list if lg.type == i.type and lg.matid == i.matid
                               and lg.goodsid == i.goodsid)
            prop = self.nm.properties(i.matid)
            goods = self.bs.tngoods(i.goodsid)
            name = '{0} {1} {2}'.format(goods.groupname, goods.name, prop.name)
            name = ' '.join(OrderedDict((w, 0) for w in name.split()).keys())
            val = (name, '', '', '', '', i.length, prop.price, self.cost())
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
        pass

    def tab_acc(self):
        acc = self.nm.acc_by_uid()
        acc_long = self.nm.acc_long()
        if not acc and not acc_long:
            return
        self.section_heading(self.row, 'Комплектующие', 'Заголовок 2')
        for ac in acc:
            name = ac.name
            if ac.article:
                name += ' арт.{}'.format(ac.article)
            val = (name, '', '', '', ac.unitsname, ac.cnt, ac.price, self.cost())
            self.put_val(val)
        for ac in acc_long:
            name = ac.name
            if ac.article:
                name += ' арт.{}'.format(ac.article)
            cost = '=D{0}*G{0}*F{0}'.format(self.row)
            val = (name, '', '', ac.length, ac.unitsname, ac.cnt, ac.price, cost)
            self.put_val(val)

    def tab_bands(self):
        pass


def start(base, path):
    rep = Report(base)
    rep.xl.new_sheet('Деталировка')
    rep.prep()
    rep.tab_sheets()
    rep.tab_long()
    rep.tab_profiles()
    rep.tab_acc()
    rep.tab_bands()
    rep.xl.save(path)
    os.startfile(path)


if __name__ == '__main__':
    num = 1
    fileDB = r'd:\К3\Самара\Самара черновик\{0}\{0}.mdb'.format(num)
    proj_rep_path = r'd:\К3\Самара\Самара черновик\{0}\Reports'.format(num)
    project = "Деталировка"
    file_path = os.path.join(proj_rep_path, '{}.xlsx'.format(project))
    db = k3r.db.DB()
    db.open(fileDB)
    start(db, file_path)
    db.close()

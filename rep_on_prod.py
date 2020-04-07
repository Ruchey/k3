# -*- coding: utf-8 -*-

import os
import k3r
from collections import namedtuple, OrderedDict


DSP = [128]
HDF = [37]
MDF = [64]
GLS = [48, 99]


class Doc:
    def __init__(self, db):
        self.bs = k3r.base.Base(db)
        self.ln = k3r.long.Long(db)
        self.nm = k3r.nomenclature.Nomenclature(db)
        self.pn = k3r.panel.Panel(db)
        self.pr = k3r.prof.Profile(db)
        self.bs = k3r.base.Base(db)
        self.xl = k3r.xl.Doc()
        self.row = 1

    def newsheet(self, name):
        """Создаём новый лист с именем"""

        self.xl.sheet_orient = self.xl.PORTRAIT
        self.xl.right_margin = 0.6
        self.xl.left_margin = 0.6
        self.xl.bottom_margin = 1.6
        self.xl.top_margin = 1.6
        self.xl.center_horizontally = 0
        self.xl.fontsize = 11
        self.xl.new_sheet(name)

    def put_val(self, val):
        self.row = self.xl.put_val(self.row, 1, val)

    def section_heading(self, row, name, style):
        self.row = self.xl.put_val(row, 1, name)
        self.xl.ws.merge_cells('A{0}:L{0}'.format(row))
        self.xl.style_to_range('A{0}:L{0}'.format(row), style)

    def det_heading(self, row):
        val = ('№', 'Наименование', 'Д-на', 'Ш-на', 'Кр. X', '', 'Кр. Y', '', 'Паз', 'Торец', 'Чер.', 'шт')
        self.xl.put_val(row, 1, val)
        self.xl.ws.merge_cells('E{0}:F{0}'.format(row))
        self.xl.ws.merge_cells('G{0}:H{0}'.format(row))
        self.xl.style_to_range('A{0}:L{0}'.format(row), 'Заголовок 3')
        self.xl.formatting(self.row, 1, ha='cccccccccccc', va='c')
        self.row += 1

    def get_pans(self, pans, tpp=None):
        keys = ('cpos', 'name', 'length', 'width', 'cnt', 'slots', 'butts', 'draw',
                'band_x1', 'band_x2', 'band_y1', 'band_y2', 'decor', 'mill')
        l_pans = []
        for id_pan in pans:
            Pans = namedtuple('Pans', keys)
            telems = self.bs.telems(id_pan)
            pdir = self.pn.dir(id_pan)
            p_cpos = telems.commonpos
            name = telems.name
            name += '(Г)' if self.pn.form(id_pan) > 0 else ''
            p_name = name
            p_length = self.pn.length(id_pan)
            p_width = self.pn.width(id_pan)
            p_cnt = telems.count
            slot_x = list(map('{0.beg}ш{0.width}г{0.depth}'.format,
                              self.pn.slots_x_par(id_pan)))
            slot_y = list(map('{0.beg}ш{0.width}г{0.depth}'.format,
                              self.pn.slots_y_par(id_pan)))
            if (45 < pdir <= 135) or (225 < pdir <= 315):
                slot_x, slot_y = slot_y, slot_x
            note_slot_x = "X={0}".format("; ".join(slot_x)) if slot_x else ""
            note_slot_y = "Y={0}".format("; ".join(slot_y)) if slot_y else ""
            p_slots = ' '.join([note_slot_x, note_slot_y])
            butts = self.pn.butts_is(id_pan)
            p_butts = 'Т' if butts else ''
            p_draw = 'Ч' if self.pn.curvepath(id_pan) > 0 else ''
            bands_abc = self.nm.bands_abc(tpp)
            band_x1 = self.pn.band_x1(id_pan)
            band_x2 = self.pn.band_x2(id_pan)
            band_y1 = self.pn.band_y1(id_pan)
            band_y2 = self.pn.band_y2(id_pan)
            p_band_x1 = bands_abc.get(band_x1.name, '')
            p_band_x2 = bands_abc.get(band_x2.name, '')
            p_band_y1 = bands_abc.get(band_y1.name, '')
            p_band_y2 = bands_abc.get(band_y2.name, '')
            dec_a = self.pn.decorates(id_pan, 5)
            dec_f = self.pn.decorates(id_pan, 6)
            decor = ' '.join([dec_a, dec_f])
            p_decor = decor
            mill = self.pn.milling(id_pan)
            p_mill = mill
            pans = Pans(p_cpos, p_name, p_length, p_width, p_cnt, p_slots, p_butts, p_draw,
                        p_band_x1, p_band_x2, p_band_y1, p_band_y2, p_decor, p_mill)
            l_pans.append(pans)
        new_list = k3r.utils.group_by_keys(l_pans, ['length', 'width', 'cnt', 'slots', 'butts', 'draw', 'band_x1',
                                                    'band_x2', 'band_y1', 'band_y2', 'decor', 'mill'], 'cnt', 'cpos')
        new_list.sort(key=lambda x: (x.length, x.width), reverse=True)
        return new_list


class Product:
    """Класс создаёт страницы деталировок по изделиям
    Входная функция make(tpp). tpp - TopParentPos это id изделия
    """
    def __init__(self, doc):
        self.doc = doc
        self.tpp = None
        self.doc.row = 1

    def make(self, tpp):
        self.tpp = tpp
        te = self.doc.bs.telems(self.tpp)
        cnt = '-- {}шт'.format(te.count)
        gab = 'в{0} ш{1} г{2}'.format(int(te.zunit), int(te.xunit), int(te.yunit))
        name = ' '.join([te.name, gab, cnt])
        self.doc.newsheet('Изд_{}'.format(self.tpp))
        cs = [4.86, 33, 7, 7, 2.71, 2.71, 2.71, 2.71, 10.14, 6, 6, 4.29]
        self.doc.xl.col_size(1, cs)
        self.doc.put_val(name)
        self.doc.xl.ws.merge_cells('A{0}:L{0}'.format(self.doc.row - 1))
        self.doc.xl.formatting(self.doc.row - 1, 1, ha='l', va='c', b='t', wrap='f', sz=14)
        self.sheets()

    def sheets(self):
        """Таблица листовых материалов"""
        mat_id = self.doc.nm.mat_by_uid(2, self.tpp)
        if not mat_id:
            return
        dsp, mdf, hdf, gls, rest = [], [], [], [], []
        lst_for_del = []
        fcd = {}
        tb_space = 8
        for i in mat_id:
            prop = self.doc.nm.properties(i)
            if prop.mattypeid in DSP:
                dsp.append(prop)
            elif prop.mattypeid in HDF:
                hdf.append(prop)
            elif prop.mattypeid in MDF:
                mdf.append(prop)
            elif prop.mattypeid in GLS:
                gls.append(prop)
            else:
                rest.append(prop)
        for group in [dsp, mdf, hdf, gls, rest]:
            if group:
                group.sort(key=lambda x: x.thickness, reverse=True)
        self.doc.det_heading(self.doc.row)
        style = ['Заголовок 62', 'Заголовок 52']
        for idx, mat_gr in enumerate([dsp, mdf]):
            for mat in mat_gr:
                l_pans = self.doc.pn.list_panels(mat.id, self.tpp)
                for i, val in enumerate(l_pans):
                    is_door = self.doc.bs.get_anc_furntype(val, '50')
                    if is_door:
                        lst_for_del.append(val)
                        is_frame = self.doc.bs.get_anc_furntype(val, '62')
                        if not is_frame:
                            if mat.id in fcd.keys():
                                fcd[mat.id].append(val)
                            else:
                                fcd.update({mat.id: [val]})
                if lst_for_del:
                    for del_el in lst_for_del:
                        l_pans.remove(del_el)
                    lst_for_del = []
                if not l_pans:
                    continue
                pans = self.doc.get_pans(l_pans)
                self.doc.section_heading(self.doc.row, mat.name, style[idx])
                self.doc.xl.style_to_range('A{0}:L{1}'.format(self.doc.row, self.doc.row + len(pans) - 1), 'Таблица 1')
                for p in pans:
                    val = (p.cpos, p.name, p.length, p.width, p.band_x1, p.band_x2, p.band_y1, p.band_y2,
                           p.slots, p.butts, p.draw, p.cnt)
                    self.doc.put_val(val)
                    self.doc.xl.formatting(self.doc.row - 1, 1, ha='llrrcccccccc', va='c', wrap='fffffffftf')
                    self.doc.xl.paint_cells('C{}'.format(self.doc.row - 1), fill='DCE6F1')
                    self.doc.xl.paint_cells('E{0}:F{0}'.format(self.doc.row - 1), fill='DCE6F1')
                self.doc.xl.row_size(self.doc.row, tb_space)
                self.doc.row += 1
        style = ['Заголовок 32', 'Заголовок 42', 'Заголовок 42']
        for idx, mat_gr in enumerate([hdf, gls, rest]):
            for mat in mat_gr:
                l_pans = self.doc.pn.list_panels(mat.id, self.tpp)
                pans = self.doc.get_pans(l_pans)
                self.doc.section_heading(self.doc.row, mat.name, style[idx])
                self.doc.xl.style_to_range('A{0}:L{1}'.format(self.doc.row, self.doc.row + len(pans) - 1), 'Таблица 1')
                for p in pans:
                    val = (p.cpos, p.name, p.length, p.width, p.band_x1, p.band_x2, p.band_y1, p.band_y2,
                           p.slots, p.butts, p.draw, p.cnt)
                    self.doc.put_val(val)
                    self.doc.xl.formatting(self.doc.row - 1, 1, ha='llrrcccccccc', va='c', wrap='fffffffftf')
                self.doc.xl.row_size(self.doc.row, tb_space)
                self.doc.row += 1
        if fcd:
            self.doc.section_heading(self.doc.row, 'Фасады', 'Заголовок 1')
            self.doc.xl.formatting(self.doc.row - 1, 1, ha='l', va='c', sz=14, b='t')
            for i in fcd:
                mat = self.doc.nm.properties(i)
                l_pans = fcd[i]
                pans = self.doc.get_pans(l_pans)
                self.doc.section_heading(self.doc.row, mat.name, 'Заголовок 42')
                self.doc.xl.style_to_range('A{0}:L{1}'.format(self.doc.row, self.doc.row + len(pans) - 1), 'Таблица 1')
                for p in pans:
                    name = ' '.join([p.name, p.decor, p.mill])
                    val = (p.cpos, name, p.length, p.width, p.band_x1, p.band_x2, p.band_y1, p.band_y2,
                           p.slots, p.butts, p.draw, p.cnt)
                    self.doc.put_val(val)
                    self.doc.xl.formatting(self.doc.row - 1, 1, ha='llrrcccccccc', va='c', wrap='fffffffftf')
            self.doc.xl.row_size(self.doc.row, tb_space)
            self.doc.row += 1
        facades = self.doc.bs.get_child_furntype(self.tpp, '50')
        fcd_ram = []
        if facades:
            for fc in facades:
                if self.doc.bs.get_child_furntype(fc[0], '62', top=0):
                    fcd_ram.append(fc)
        if fcd_ram:
            self.doc.section_heading(self.doc.row, 'Фасады рамочные', 'Заголовок 1')
            self.doc.xl.formatting(self.doc.row - 1, 1, ha='l', va='c', sz=14, b='t')
            self.doc.xl.style_to_range('A{0}:L{1}'.format(self.doc.row, self.doc.row + len(fcd_ram) - 1), 'Таблица 1')
            for i in fcd_ram:
                te = self.doc.bs.telems(i[0])
                cpos = te.commonpos
                name = te.name
                height = te.zunit
                width = te.xunit
                cnt = te.count
                val = (cpos, name, height, width, '', '', '', '', '', '', '', cnt)
                self.doc.put_val(val)
                self.doc.xl.formatting(self.doc.row - 1, 1, ha='llrrcccccccc', va='c', wrap='fffffffftf')
            self.doc.xl.row_size(self.doc.row, tb_space)
            self.doc.row += 1
        bands_abc = self.doc.nm.bands_abc(self.tpp)
        if bands_abc:
            self.doc.section_heading(self.doc.row, 'Расшифровка кромки', 'Заголовок 6')
            self.doc.xl.style_to_range('A{0}:L{1}'.format(self.doc.row, self.doc.row + len(bands_abc) - 1), 'Линейка 1')
            for i in bands_abc:
                val = (bands_abc[i], i)
                self.doc.put_val(val)


class Detailing:
    def __init__(self, doc):
        self.doc = doc
        self.doc.row = 1

    def make(self):
        self.doc.newsheet('Деталировка')
        cs = [4.86, 33, 7, 7, 2.71, 2.71, 2.71, 2.71, 10.14, 6, 6, 4.29]
        self.doc.xl.col_size(1, cs)
        self.detailing()

    def detailing(self):
        pr = Product(self.doc)
        pr.sheets()


class Report:
    def __init__(self, db):
        self.doc = Doc(db)

    def make(self):
        prod = Product(self.doc)
        objects = self.doc.bs.tobjects()
        # for obj in objects:
        #     prod.make(obj.unitpos)
        det = Detailing(self.doc)
        det.make()

    def save(self, file):
        self.doc.xl.save(file)


def start(fileDB, projreppath, project):
    # Подключаемся к базе выгрузки
    db = k3r.db.DB()
    db.open(fileDB)

    rep = Report(db)
    rep.make()
    db.close()
    file = os.path.join(projreppath, '{}.xlsx'.format(project))
    rep.save(file)
    os.startfile(file)
    return True


if __name__ == '__main__':
    fileDB = r'd:\К3\Самара\Самара черновик\2\2.mdb'
    projreppath = r'd:\К3\Самара\Самара черновик\2\Reports'
    project = "Общий отчёт"
    start(fileDB, projreppath, project)

"""
Отчёт для К3-Мебель.
Выводит листы:
    Общая деталировка
    Деталировка на каждое изделие
    Спецификация
    Лист раскроя профилей

Используемые свойства, которых нет в стандартной настройке К3-мебель:
    density - тип Число - плотность материала. Для ЛДСП 780 кг/м3
    supplier - тип Строка - поставщик фурнитуры или материалов
    notcutpc - тип Да/Нет - не резать на куски. Используется в профилях, которые не надо выводить кусками
    stepcut - тип Размер - шаг нарезки профиля

"""

import math
import k3r
from collections import namedtuple
from pathlib import Path


class Doc:
    def __init__(self, db, **kwargs):
        self.bs = k3r.base.Base(db)
        self.ln = k3r.long.Long(db)
        self.nm = k3r.nomenclature.Nomenclature(db)
        self.pn = k3r.panel.Panel(db)
        self.pf = k3r.prof.Profile(db)
        self.bs = k3r.base.Base(db)
        self.xl = k3r.xl.Doc()
        self.gt = k3r.get_tables.Specific(db)
        self.row = 1

    def new_sheet(self, name, tab_color=None):
        """Создаём новый лист с именем"""

        self.xl.sheet_orient = self.xl.PORTRAIT
        self.xl.right_margin = 0.6
        self.xl.left_margin = 0.6
        self.xl.bottom_margin = 1.6
        self.xl.top_margin = 1.6
        self.xl.center_horizontally = 0
        self.xl.fontsize = 11
        self.xl.new_sheet(name, tab_color)

    def cap(self, pg=2):
        """Шапка с данными заказа"""
        to = self.bs.torderinfo()
        number = to.ordernumber
        name = to.ordername if to.ordername else ""
        customer = to.customer if to.customer else ""
        phone = to.telephonenumber if to.telephonenumber else ""
        adr = to.address if to.address else ""
        executor = to.executor if to.executor else ""
        acceptor = to.acceptor if to.acceptor else ""
        addinfo = to.additionalinfo if to.additionalinfo else ""
        data = ""
        if to.orderexpirationdata.year > 2000:
            data = to.orderexpirationdata.strftime("%m.%d.%y")
        val1 = [
            ("Заказ №{0}".format(number)),
            ("Название: {0}".format(name)),
            ("Заказчик: {0} Телефон: {1}".format(customer, phone)),
            ("Адрес: {0}".format(adr)),
            ("Приёмщик: {0} Исполнитель: {1}".format(acceptor, executor)),
            ("Дата исполнения: {}".format(data)),
            ("Дополнительно: {}".format(addinfo)),
        ]
        val2 = [
            ("Заказ №{0} {1}".format(number, name)),
            ("Дата исполнения: {}".format(data)),
        ]
        val = {1: val1, 2: val2}
        for v in val[pg]:
            self.xl.formatting(self.row, 1, sz=12, bld="f", itl="f")
            self.put_val(v)
        self.xl.row_size(self.row, 8)
        self.row += 1

    def put_val(self, val):
        self.row = self.xl.put_val(self.row, 1, val)

    def section_heading(self, row, name, style, end_col="L"):
        self.row = self.xl.put_val(row, 1, name)
        self.xl.ws.merge_cells("A{0}:{1}{0}".format(row, end_col))
        self.xl.style_to_range("A{0}:{1}{0}".format(row, end_col), style)

    def header(self, val_1, val_2, tab=None):
        self.xl.put_val(self.row, 1, val_1[0].upper() + "    ")
        self.xl.style_to_range("A{0}:E{0}".format(self.row), val_1[1])
        self.xl.ws.merge_cells("A{0}:E{0}".format(self.row))
        self.xl.formatting(self.row, 1, ha="r", va="c", wrap="f", sz=12, bld="t")
        self.row += 1
        self.xl.put_val(self.row, 1, val_2[0])
        self.xl.style_to_range("A{0}:E{0}".format(self.row), val_2[1])
        self.xl.formatting(self.row, 1, ha="ccccc", va="c", wrap="f")
        self.row += 1
        if tab:
            self.xl.style_to_range(
                "A{0}:E{1}".format(self.row, self.row + tab), "Таблица 1"
            )

    def det_heading(self, row, billet=False):
        val = (
            "№",
            "Наименование",
            "Деталь",
            "",
            "Кромка",
            "",
            "",
            "",
            "Паз X-отст. Y-отст.",
            "Т",
            "Ч",
            "шт",
        )
        val2 = ("X", "Y", "X1", "X2", "Y1", "Y2")
        val_billet = (
            "№",
            "Наименование",
            "Заготовка",
            "",
            "Деталь",
            "",
            "Кромка",
            "",
            "",
            "",
            "Паз X-отст. Y-отст.",
            "Т",
            "Ч",
            "шт",
        )
        val2_billet = ("X", "Y", "X", "Y", "X1", "X2", "Y1", "Y2")
        if billet:
            self.xl.put_val(row, 1, val_billet)
            self.xl.put_val(row+1, 3, val2_billet)
            self.xl.ws.merge_cells("C{0}:D{0}".format(row))
            self.xl.ws.merge_cells("E{0}:F{0}".format(row))
            self.xl.ws.merge_cells("G{0}:J{0}".format(row))
            self.xl.style_to_range("A{0}:N{1}".format(row, row+1), "Заголовок 3")
            self.xl.formatting(self.row, 1, ha="cccccccccccccc", va="c", wrap="t")
            self.xl.formatting(self.row+1, 1, ha="cccccccccccccc", va="c", wrap="t")
        else:
            self.xl.put_val(row, 1, val)
            self.xl.put_val(row+1, 3, val2)
            self.xl.ws.merge_cells("C{0}:D{0}".format(row))
            self.xl.ws.merge_cells("E{0}:H{0}".format(row))
            self.xl.style_to_range("A{0}:L{1}".format(row, row+1), "Заголовок 3")
            self.xl.formatting(self.row, 1, ha="cccccccccccc", va="c", wrap="t")
            self.xl.formatting(self.row+1, 1, ha="cccccccccccc", va="c", wrap="t")

        self.row += 2

    def get_pans(self, pans, tpp=None):
        """Входные данные:
        pans - писок id панелей
        tpp - TopParentPos родитель
        fugue - размер прифуговки для кромления (default 1mm)
        """
        keys = (
            "cpos",
            "name",
            "plane_length",
            "plane_width",
            "length",
            "width",
            "cnt",
            "slots",
            "butts",
            "draw",
            "band_x1",
            "band_x2",
            "band_y1",
            "band_y2",
            "decor",
            "mill",
        )
        l_pans = []
        for id_pan in pans:
            Pans = namedtuple("Pans", keys)
            telems = self.bs.telems(id_pan)
            pdir = self.pn.dir(id_pan)
            p_cpos = telems.commonpos
            name = telems.name
            plane_length = self.pn.planelength(id_pan, self.fugue)
            plane_width = self.pn.planewidth(id_pan, self.fugue)
            p_length = self.pn.length(id_pan)
            p_width = self.pn.width(id_pan)
            if self.pn.form(id_pan) > 0:
                par_bent = self.pn.par_bent_pan(id_pan)
                ax = par_bent.ax
                width = p_width if ax == 1 else p_length
                name += " R{0} (разв. {1})".format(par_bent.rad, round(width))
                p_width = par_bent.chord
                p_length = par_bent.h
            p_name = name
            p_cnt = telems.count
            slot_x = list(
                map("{0.beg}ш{0.width}г{0.depth}".format, self.pn.slots_x_par(id_pan))
            )
            slot_y = list(
                map("{0.beg}ш{0.width}г{0.depth}".format, self.pn.slots_y_par(id_pan))
            )
            if (45 < pdir <= 135) or (225 < pdir <= 315):
                slot_x, slot_y = slot_y, slot_x
            note_slot_x = "X_{0}".format("; ".join(slot_x)) if slot_x else ""
            note_slot_y = "Y_{0}".format("; ".join(slot_y)) if slot_y else ""
            p_slots = " ".join([note_slot_x, note_slot_y])
            butts = self.pn.butts_is(id_pan)
            p_butts = "Т" if butts else ""
            p_draw = "Ч" if self.pn.curvepath(id_pan) > 0 else ""
            bands_abc = self.nm.bands_abc(tpp)
            band_x1 = self.pn.band_x1(id_pan)
            band_x2 = self.pn.band_x2(id_pan)
            band_y1 = self.pn.band_y1(id_pan)
            band_y2 = self.pn.band_y2(id_pan)
            p_band_x1 = bands_abc.get(band_x1.name, "")
            p_band_x2 = bands_abc.get(band_x2.name, "")
            p_band_y1 = bands_abc.get(band_y1.name, "")
            p_band_y2 = bands_abc.get(band_y2.name, "")
            dec_a = self.pn.decorates(id_pan, 5)
            dec_f = self.pn.decorates(id_pan, 6)
            decor = " ".join([dec_a, dec_f]).strip()
            p_decor = decor
            mill = self.pn.milling(id_pan)
            p_mill = mill
            pans = Pans(
                p_cpos,
                p_name,
                plane_length,
                plane_width,
                p_length,
                p_width,
                p_cnt,
                p_slots,
                p_butts,
                p_draw,
                p_band_x1,
                p_band_x2,
                p_band_y1,
                p_band_y2,
                p_decor,
                p_mill,
            )
            l_pans.append(pans)
        new_list = k3r.utils.group_by_keys(
            l_pans,
            (
                "length",
                "width",
                "cnt",
                "slots",
                "butts",
                "draw",
                "band_x1",
                "band_x2",
                "band_y1",
                "band_y2",
                "decor",
                "mill",
            ),
            "cnt",
            "cpos",
        )
        new_list.sort(key=lambda x: (x.length, x.width), reverse=True)
        return new_list


class Product:
    """Класс создаёт страницы деталировок по изделиям
    Входная функция make(tpp). tpp - TopParentPos это id изделия
    """

    def __init__(self, doc, **kwargs):
        """Входные данные:
        doc - документ
        billet - заготовка (выводить размеры заготовки)
        """
        self.doc = doc
        self.doc.row = 1
        self.tpp = None
        self.billet = kwargs.get("billet", False)
        self.end_col = "L" if not self.billet else "N"

    def make(self, tpp):
        self.tpp = tpp
        te = self.doc.bs.telems(self.tpp)
        t_obj = self.doc.bs.tobjects(self.tpp)
        pt = t_obj.placetype
        sheet_name = "низ_{}".format(self.tpp)
        color = "9FC0E7"
        if pt == 1:
            sheet_name = "верх_{}".format(self.tpp)
            color = "FCDABC"
        self.doc.new_sheet(sheet_name, tab_color=color)
        self.doc.cap()
        column_size = [5.29, 43.14, 7, 7, 2.71, 2.71, 2.71, 2.71, 10, 1.29, 1.29, 4.29]
        column_size_billet = [4.86, 33, 4.57, 4.57, 2.71, 2.71, 2.71, 2.71, 10, 1.29, 1.29, 4.29]
        cs = column_size_billet if self.billet else column_size
        self.doc.xl.col_size(1, cs)
        gab = "в{0} ш{1} г{2}".format(int(te.zunit), int(te.xunit), int(te.yunit))
        cnt = "-- {}шт".format(te.count)
        name = " ".join([te.name, gab, cnt])
        self.doc.put_val(name)
        self.doc.xl.ws.merge_cells("A{0}:{1}{0}".format(self.doc.row - 1, self.end_col))
        self.doc.xl.formatting(
            self.doc.row - 1, 1, ha="c", va="c", wrap="f", sz=14, bld="t"
        )
        self.sheets()
        self.doc.row = 1

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
            if prop.mattypeid in k3r.variables.DSP:
                dsp.append(prop)
            elif prop.mattypeid in k3r.variables.HDF:
                hdf.append(prop)
            elif prop.mattypeid in k3r.variables.MDF:
                mdf.append(prop)
            elif prop.mattypeid in k3r.variables.GLS:
                gls.append(prop)
            else:
                rest.append(prop)
        for group in [dsp, mdf, hdf, gls, rest]:
            if group:
                group.sort(key=lambda x: x.thickness, reverse=True)
        bands_abc = self.doc.nm.bands_abc(self.tpp)
        if bands_abc:
            self.doc.section_heading(self.doc.row, "Расшифровка кромки", "Заголовок 6", end_col=self.end_col)
            self.doc.xl.style_to_range(
                "A{0}:{2}{1}".format(self.doc.row, self.doc.row + len(bands_abc) - 1, self.end_col),
                "Линейка 1",
            )
            for i in bands_abc:
                val = (bands_abc[i], i)
                self.doc.put_val(val)
        self.doc.det_heading(self.doc.row, billet=self.billet)
        style = ["Заголовок 62", "Заголовок 52"]
        for idx, mat_gr in enumerate([dsp, mdf]):
            for mat in mat_gr:
                l_pans = self.doc.pn.list_panels(mat.priceid, self.tpp)
                for i, val in enumerate(l_pans):
                    is_door = self.doc.bs.get_anc_furntype(val, "50")
                    dec_a = self.doc.pn.decorates(val, 5)
                    dec_f = self.doc.pn.decorates(val, 6)
                    if is_door or dec_a or dec_f:
                        lst_for_del.append(val)
                        is_frame = self.doc.bs.get_anc_furntype(val, "62")
                        if not is_frame:
                            if mat.priceid in fcd.keys():
                                fcd[mat.priceid].append(val)
                            else:
                                fcd.update({mat.priceid: [val]})
                if lst_for_del:
                    for del_el in lst_for_del:
                        l_pans.remove(del_el)
                    lst_for_del = []
                if not l_pans:
                    continue
                pans = self.doc.get_pans(l_pans)
                self.doc.section_heading(self.doc.row, mat.name, style[idx], end_col=self.end_col)
                self.doc.xl.style_to_range(
                    "A{0}:{2}{1}".format(self.doc.row, self.doc.row + len(pans) - 1, self.end_col),
                    "Таблица 1",
                )
                for p in pans:
                    if self.billet:
                        val = (
                            p.cpos,
                            p.name,
                            p.plane_length,
                            p.plane_width,
                            p.length,
                            p.width,
                            p.band_x1,
                            p.band_x2,
                            p.band_y1,
                            p.band_y2,
                            p.slots,
                            p.butts,
                            p.draw,
                            p.cnt,
                        )
                        self.doc.put_val(val)
                        self.doc.xl.formatting(
                            self.doc.row - 1,
                            1,
                            ha="llrrrrcccccccc",
                            va="c",
                            wrap="tffffffffftfff",
                        )
                        self.doc.xl.paint_cells(
                            "E{}".format(self.doc.row - 1), fill="DCE6F1"
                        )
                        self.doc.xl.paint_cells(
                            "G{0}:H{0}".format(self.doc.row - 1), fill="DCE6F1"
                        )
                    else:
                        val = (
                            p.cpos,
                            p.name,
                            p.length,
                            p.width,
                            p.band_x1,
                            p.band_x2,
                            p.band_y1,
                            p.band_y2,
                            p.slots,
                            p.butts,
                            p.draw,
                            p.cnt,
                        )
                        self.doc.put_val(val)
                        self.doc.xl.formatting(
                            self.doc.row - 1,
                            1,
                            ha="llrrcccccccc",
                            va="c",
                            wrap="tffffffftfff",
                            )
                        self.doc.xl.paint_cells(
                            "C{}".format(self.doc.row - 1), fill="DCE6F1"
                        )
                        self.doc.xl.paint_cells(
                            "E{0}:F{0}".format(self.doc.row - 1), fill="DCE6F1"
                        )
                self.doc.xl.row_size(self.doc.row, tb_space)
                self.doc.row += 1
        style = ["Заголовок 32", "Заголовок 42", "Заголовок 42"]
        for idx, mat_gr in enumerate([hdf, gls, rest]):
            for mat in mat_gr:
                l_pans = self.doc.pn.list_panels(mat.priceid, self.tpp)
                pans = self.doc.get_pans(l_pans)
                self.doc.section_heading(self.doc.row, mat.name, style[idx], end_col=self.end_col)
                self.doc.xl.style_to_range(
                    "A{0}:{2}{1}".format(self.doc.row, self.doc.row + len(pans) - 1, self.end_col),
                    "Таблица 1",
                )
                for p in pans:
                    if self.billet:
                        val = (
                            p.cpos,
                            p.name,
                            "",
                            "",
                            p.length,
                            p.width,
                            p.band_x1,
                            p.band_x2,
                            p.band_y1,
                            p.band_y2,
                            p.slots,
                            p.butts,
                            p.draw,
                            p.cnt,
                        )
                        self.doc.put_val(val)
                        self.doc.xl.formatting(
                            self.doc.row - 1,
                            1,
                            ha="llrrcccccccccc",
                            va="c",
                            wrap="fffffffftfff",
                        )
                    else:
                        val = (
                            p.cpos,
                            p.name,
                            p.length,
                            p.width,
                            p.band_x1,
                            p.band_x2,
                            p.band_y1,
                            p.band_y2,
                            p.slots,
                            p.butts,
                            p.draw,
                            p.cnt,
                        )
                        self.doc.put_val(val)
                        self.doc.xl.formatting(
                            self.doc.row - 1,
                            1,
                            ha="llrrcccccccc",
                            va="c",
                            wrap="fffffftfff",
                            )
                self.doc.xl.row_size(self.doc.row, tb_space)
                self.doc.row += 1
        if fcd:
            self.doc.section_heading(self.doc.row, "Фасады", "Заголовок 1", end_col=self.end_col)
            self.doc.xl.formatting(self.doc.row - 1, 1, ha="l", va="c", sz=14, b="t")
            for i in fcd:
                mat = self.doc.nm.properties(i)
                l_pans = fcd[i]
                pans = self.doc.get_pans(l_pans)
                gr_pans = k3r.utils.group_by_keys(pans, ("decor", "mill"))
                for gr in gr_pans:
                    name = " ".join([mat.name, gr[0].decor, gr[0].mill])
                    self.doc.section_heading(self.doc.row, name, "Заголовок 42", end_col=self.end_col)
                    self.doc.xl.style_to_range(
                        "A{0}:{2}{1}".format(self.doc.row, self.doc.row + len(gr) - 1, self.end_col),
                        "Таблица 1",
                    )
                    for p in gr:
                        if self.billet:
                            val = (
                                p.cpos,
                                p.name,
                                p.plane_length,
                                p.plane_width,
                                p.length,
                                p.width,
                                p.band_x1,
                                p.band_x2,
                                p.band_y1,
                                p.band_y2,
                                p.slots,
                                p.butts,
                                p.draw,
                                p.cnt,
                            )
                            self.doc.put_val(val)
                            self.doc.xl.formatting(
                                self.doc.row - 1,
                                1,
                                ha="llrrcccccccccc",
                                va="c",
                                wrap="fffffffftfff",
                            )
                        else:
                            val = (
                                p.cpos,
                                p.name,
                                p.length,
                                p.width,
                                p.band_x1,
                                p.band_x2,
                                p.band_y1,
                                p.band_y2,
                                p.slots,
                                p.butts,
                                p.draw,
                                p.cnt,
                            )
                            self.doc.put_val(val)
                            self.doc.xl.formatting(
                                self.doc.row - 1,
                                1,
                                ha="llrrcccccccc",
                                va="c",
                                wrap="fffffftfff",
                                )
            self.doc.xl.row_size(self.doc.row, tb_space)
            self.doc.row += 1
        facades = self.doc.bs.get_child_furntype(self.tpp, "50")
        fcd_ram = []
        if facades:
            for fc in facades:
                if self.doc.bs.get_child_furntype(fc[0], "62", top=0):
                    fcd_ram.append(fc)
        if fcd_ram:
            self.doc.section_heading(self.doc.row, "Фасады рамочные", "Заголовок 1", end_col=self.end_col)
            self.doc.xl.formatting(self.doc.row - 1, 1, ha="l", va="c", sz=14, b="t")
            self.doc.xl.style_to_range(
                "A{0}:{2}{1}".format(self.doc.row, self.doc.row + len(fcd_ram) - 1, self.end_col),
                "Таблица 1",
            )
            for i in fcd_ram:
                te = self.doc.bs.telems(i[0])
                cpos = te.commonpos
                name = te.name
                height = te.zunit
                width = te.xunit
                cnt = te.count
                val = (cpos, name, height, width, "", "", "", "", "", "", "", cnt)
                self.doc.put_val(val)
                self.doc.xl.formatting(
                    self.doc.row - 1, 1, ha="llrrcccccccc", va="c", wrap="fffffffftf"
                )
            self.doc.xl.row_size(self.doc.row, tb_space)
            self.doc.row += 1


class Detailing:
    def __init__(self, doc, **kwargs):
        """Входные данные:
        doc - документ
        billet - заготовка (выводить размеры заготовки)
        """
        self.doc = doc
        self.doc.row = 1
        self.billet = kwargs.get("billet", False)
        self.end_col = "L" if not self.billet else "N"

    def make(self):
        self.doc.new_sheet("Деталировка", tab_color="595959")
        column_size = [5.29, 43.14, 7, 7, 2.71, 2.71, 2.71, 2.71, 10, 1.29, 1.29, 4.29]
        column_size_billet = [4.86, 32.57, 6.29, 6.29, 6.29, 6.29, 2.71, 2.71, 2.71, 2.71, 10, 1.29, 1.29, 3.29]
        cs = column_size_billet if self.billet else column_size
        self.doc.xl.col_size(1, cs)
        self.detailing()

    def detailing(self):
        product = Product(self.doc, billet=self.billet)
        self.doc.cap()
        product.sheets()


class Specification:
    def __init__(self, doc):
        self.doc = doc
        self.doc.row = 1
        self.tpp = None

    def make(self, tpp=None):
        self.tpp = tpp
        self.doc.new_sheet("Спецификация", tab_color="FFFF00")
        column_size = [58, 12.71, 12, 6.86, 6]
        self.doc.xl.col_size(1, column_size)
        self.doc.cap()
        self.sheets()
        self.glass()
        self.bands()
        self.prof()
        self.acc()

    def sheets(self):
        sh = self.doc.gt.t_sheets(self.tpp)
        if not sh:
            return
        val_1 = ("Листовой материал", "Заголовок 32")
        val_2 = (("Материал панелей", "Артикул", "Формат", "кв.м", "листов"), "Заголовок 1")
        self.doc.header(val_1, val_2, len(sh) - 1)
        for i in sh:
            #density = getattr(i, "density", 0)
            gabx = int(getattr(i, "gabx", 1))
            gaby = int(getattr(i, "gaby", 1))
            #weight = round(i.sqm * density * (i.thickness / 1000), 1)
            sheet_size = "{0}x{1}x{2}".format(gabx, gaby, int(i.thickness))
            number_sheets = round((i.sqm * i.wastecoeff) / (gabx * gaby / 1000000), 1)
            val = (i.name, i.article, sheet_size, i.sqm, number_sheets)
            self.doc.put_val(val)
            self.doc.xl.formatting(self.doc.row - 1, 1, ha="lrrrr", va="c")
        self.doc.xl.row_size(self.doc.row, 8)
        self.doc.row += 1

    def glass(self):
        gs = self.doc.gt.t_glass(self.tpp)
        if not gs:
            return
        val_1 = ("Стёкла и зеркала", "Заголовок 72")
        val_2 = (("Материал", "Длина", "Ширина", "ед.изм", "кол-во"), "Заголовок 1")
        self.doc.header(val_1, val_2, 0)
        for i in gs:
            l_pans = self.doc.pn.list_panels(i.priceid, self.tpp)
            if not l_pans:
                continue
            pans = self.doc.get_pans(l_pans)
            val = (i.name, "", "", i.unitsname, i.sqm)
            self.doc.put_val(val)
            self.doc.xl.style_to_range(
                "A{0}:E{0}".format(self.doc.row - 1), "Заголовок 8"
            )
            self.doc.xl.formatting(self.doc.row - 1, 1, ha="lrrcr", va="c", bld="t")
            for p in pans:
                name = p.name
                bnd = "x".join(
                    filter(bool, [p.band_x1, p.band_x2, p.band_y1, p.band_y2])
                )
                if bnd:
                    name += " обработка {}".format(bnd)
                val = (name, p.length, p.width, "шт", p.cnt)
                self.doc.put_val(val)
                self.doc.xl.formatting(self.doc.row - 1, 1, ha="lrrcr", va="c")

    def bands(self):
        bnd = self.doc.gt.t_bands(tpp=self.tpp, by_thick=False)
        if not bnd:
            return
        val_1 = ("Кромочный материал", "Заголовок 42")
        val_2 = (
            ("Наименование кромочного материала", "", "Артикул", "", "п/м"),
            "Заголовок 1",
        )
        self.doc.header(val_1, val_2, len(bnd) - 1)
        bands_abc = self.doc.nm.bands_abc(self.tpp)
        for i in bnd:
            name = "{0} | {1}".format(bands_abc[i.name], i.name)
            val = (name, i.article, "", "", math.ceil(i.len * i.wastecoeff))
            self.doc.put_val(val)
            self.doc.xl.ws.merge_cells("B{0}:D{0}".format(self.doc.row - 1))
            self.doc.xl.formatting(self.doc.row - 1, 1, ha="lcccr", va="c")
        self.doc.xl.row_size(self.doc.row, 8)
        self.doc.row += 1

    def prof(self):
        pf = self.doc.gt.t_total_prof(tpp=self.tpp)
        if not pf:
            return
        val_1 = ("Профиля", "Заголовок 52")
        val_2 = (("Наименование профиля", "", "Артикул", "", "п/м"), "Заголовок 1")
        self.doc.header(val_1, val_2, len(pf) - 1)
        for i in pf:
            val = (i.name, "", i.article, "", round(i.len, 2))
            self.doc.put_val(val)
            self.doc.xl.ws.merge_cells("A{0}:B{0}".format(self.doc.row - 1))
            self.doc.xl.ws.merge_cells("C{0}:D{0}".format(self.doc.row - 1))
            self.doc.xl.formatting(self.doc.row - 1, 1, ha="llccr", va="c")
        self.doc.xl.row_size(self.doc.row, 8)
        self.doc.row += 1

    def acc(self):
        ac = self.doc.gt.t_acc(tpp=self.tpp)
        if not ac:
            return
        val_1 = ("Комплектующие", "Заголовок 62")
        val_2 = (
            ("Наименование", "Артикул", "Поставщик", "ед.изм", "кол-во"),
            "Заголовок 1",
        )
        self.doc.header(val_1, val_2, len(ac) - 1)
        for i in ac:
            supp = getattr(i, "supplier", "")
            val = (i.name, i.article, supp, i.unitsname, i.cnt)
            self.doc.put_val(val)
            self.doc.xl.formatting(self.doc.row - 1, 1, ha="lcclr", va="c", wrap="t")
        ac_long = self.doc.gt.t_acc_long(tpp=self.tpp)
        if not ac_long:
            self.doc.xl.row_size(self.doc.row, 8)
            self.doc.row += 1
            return
        for i in ac_long:
            begin = self.doc.row
            end = self.doc.row + len(ac_long) - 2
            self.doc.xl.style_to_range("A{0}:E{1}".format(begin, end), "Таблица 1")
            supp = getattr(i, "supplier", "")
            name = "{0} (L={1}мм)".format(i.name, int(i.len))
            val = (name, i.article, supp, i.unitsname, i.cnt)
            self.doc.put_val(val)
            self.doc.xl.formatting(self.doc.row - 1, 1, ha="lcclr", va="c", wrap="t")
        self.doc.xl.row_size(self.doc.row, 8)
        self.doc.row += 1


class Profiles:
    def __init__(self, doc):
        self.doc = doc
        self.doc.row = 1
        self.tpp = None

    def make(self, tpp=None):
        self.tpp = tpp
        if not self.doc.pf.profiles(tpp=self.tpp):
            return
        self.doc.new_sheet("Профиля", tab_color="3AE2CE")
        cs = [53, 12.71, 12, 6.86, 8.57]
        self.doc.xl.col_size(1, cs)
        self.doc.cap()
        self.prof()

    def prof(self):
        total_pf = self.doc.gt.t_total_prof(tpp=self.tpp)
        profiles = self.doc.gt.t_profiles(tpp=self.tpp)
        val_1 = ("Профиля", "Заголовок 62")
        val_2 = (("Наименование профиля", "", "Длина, мм", "", "шт"), "Заголовок 2")
        self.doc.header(val_1, val_2, 0)
        self.doc.xl.ws.merge_cells("C{0}:D{0}".format(self.doc.row - 1))
        for i in total_pf:
            val = (i.name, "", i.article, "", i.net_len)
            self.doc.put_val(val)
            self.doc.xl.style_to_range(
                "A{0}:E{0}".format(self.doc.row - 1), "Заголовок 1"
            )
            self.doc.xl.formatting(
                self.doc.row - 1,
                1,
                ha="llccr",
                va="c",
                bld="t",
                itl="t",
                nf=("@", "@", "арт. \@", "@", "0.0\ п\/\м"),
            )
            pf_by_mat = list(filter(lambda x: x.priceid == i.priceid, profiles))
            if not i.notcutpc:
                start_row = self.doc.row
                end_row = self.doc.row + len(pf_by_mat) - 1
                self.doc.xl.style_to_range(
                    "A{0}:E{1}".format(start_row, end_row), "Таблица 1"
                )
                for j in pf_by_mat:
                    val = (j.elemname, "", j.len, "", j.cnt)
                    self.doc.put_val(val)
                    self.doc.xl.ws.merge_cells("A{0}:B{0}".format(self.doc.row - 1))
                    self.doc.xl.ws.merge_cells("C{0}:D{0}".format(self.doc.row - 1))
                    self.doc.xl.formatting(
                        self.doc.row - 1,
                        1,
                        ha="llrrr",
                        va="c",
                        nf=("-- @", "@", "# мм", "# шт"),
                    )
                    self.doc.xl.paint_cells(
                        "A{0}".format(self.doc.row - 1), ink="808080"
                    )
                    self.doc.xl.ws.row_dimensions.group(
                        start_row, end_row, hidden=False
                    )


class Facades:
    def __init__(self, doc):
        self.doc = doc
        self.doc.row = 1
        self.tpp = None

    def make(self, tpp=None):
        self.tpp = tpp
        if not self.doc.pf.profiles(tpp=self.tpp):
            return
        self.doc.cap()
        self.doc.new_sheet("Фасады", tab_color="3AE2CE")
        cs = [5.29, 7.43, 43, 10.14, 10.14, 8.43]
        self.doc.xl.col_size(1, cs)
        self.facades()


class Report:
    def __init__(self, **kwargs):
        db_path = kwargs.get("db_path", "")
        fugue = kwargs.get("fugue", 1)
        self.db = k3r.db.DB(db_path)
        self.db.open()
        self.doc = Doc(self.db, fugue=fugue)

    def __del__(self):
        print("Закрываем соединение с базой отчёта")
        self.db.close()

    def make(self):
        detailing = Detailing(self.doc, billet=True)
        detailing.make()
        specification = Specification(self.doc)
        specification.make()
        profiles = Profiles(self.doc)
        profiles.make()
        product = Product(self.doc)
        objects = self.doc.bs.tobjects()
        objects.sort(key=lambda x: x.placetype)
        for obj in objects:
            product.make(obj.unitpos)

    def save(self, file):
        self.doc.xl.save(file)



def start(file_db, pr_path, name):
    if os.path.exists(file_db) == False:
        return False
    rep = Report(file_db)
    rep.make()
    file = os.path.join(pr_path, "{}.xlsx".format(name))
    rep.save(file)
    os.startfile(file)
    return True


if __name__ == "__main__":
    nm = 5
    fileDB = r"d:\К3\HL\2020\{0}\{0}.mdb".format(nm)
    pr_rep_path = r"d:\К3\HL\2020\{}\Reports".format(nm)
    rep_name = "Общий отчёт"
    start(fileDB, pr_rep_path, rep_name)

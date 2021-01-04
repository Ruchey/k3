# -*- coding: utf-8 -*-

import itertools
import math

from . import utils
from collections import namedtuple
from functools import lru_cache

__author__ = "Виноградов А.Г. г.Белгород  август 2015"


class Panel:
    def __init__(self, db):
        """Информация о панели"""
        self.db = db

    def list_panels(self, mat_id=None, tpp=None):
        """Получить список панелей по материалу
        Keyword arguments:
            mat_id - id материала из которого сделана панель
            tpp - главный родитель панели
        Returns:
            [id 1, id 2, ..., id n]
        """

        filter_id = "AND tnn.ID={}".format(mat_id) if mat_id else ""
        filter_tpp = "AND te.TopParentPos={}".format(tpp) if tpp else ""
        where = " ".join(["te.FurnType Like '01%'", filter_id, filter_tpp])

        sql = (
            "SELECT te.UnitPos FROM TElems AS te INNER JOIN TNNomenclature AS tnn "
            "ON te.PriceID = tnn.ID WHERE {0} ORDER BY te.CommonPos".format(where)
        )
        res = self.db.rs(sql)
        list_id = []
        if res:
            for i in res:
                list_id.append(i[0])
        return list_id

    @lru_cache(maxsize=6)
    def xunit(self, unitpos):
        """Размер панели вдоль X"""

        sql = "SELECT xunit FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return utils.float_int(res[0][0])
        else:
            return 0

    @lru_cache(maxsize=6)
    def yunit(self, unitpos):
        """Размер панели вдоль Y"""

        sql = "SELECT yunit FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return utils.float_int(res[0][0])
        else:
            return 0

    @lru_cache(maxsize=6)
    def zunit(self, unitpos):
        """Размер панели вдоль Z"""

        sql = "SELECT zunit FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return utils.float_int(res[0][0])
        else:
            return 0

    def cnt(self, unitpos):
        """Количество панелей. Возвращает св-во count из таблицы TElems"""

        sql = "SELECT count FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return res[0][0]

    def length(self, unitpos):
        """Длина панели с кромкой"""

        sql = "SELECT length FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return utils.float_int(res[0][0])
        else:
            return 0

    def width(self, unitpos):
        """Ширина панели с кромкой"""

        sql = "SELECT width FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return utils.float_int(res[0][0])
        else:
            return 0

    def planelength(self, unitpos, fugue=0):
        """Длина панели без кромки
        unitpos - int уникальный номер панели
        fugue - int|float размер прифуговки на торец с кромкой (default 0)
        Если панели с таким номером нет, вернёт ноль.
        """

        sql = "SELECT planelength FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            band_y1 = self.band_y1(unitpos)
            band_y2 = self.band_y2(unitpos)
            add_size = (int(bool(band_y1.id)) + int(bool(band_y2.id))) * fugue
            return utils.float_int(res[0][0]+add_size)
        else:
            return 0

    def planewidth(self, unitpos, fugue=0):
        """Ширина панели без кромки
        unitpos - int уникальный номер панели
        fugue - int|float размер прифуговки на торец с кромкой (default 0)
        Если панели с таким номером нет, вернёт ноль.
        """

        sql = "SELECT planewidth FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            band_x1 = self.band_x1(unitpos)
            band_x2 = self.band_x2(unitpos)
            add_size = (int(bool(band_x1.id)) + int(bool(band_x2.id))) * fugue
            return utils.float_int(res[0][0]+add_size)
        else:
            return 0

    def thickness(self, unitpos):
        """Толщина панели"""

        sql = "SELECT thickness FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return utils.float_int(res[0][0])
        else:
            return 0

    @lru_cache(maxsize=6)
    def dir(self, unitpos):
        """Направление текстуры в градусах"""

        sql = "SELECT dir FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return utils.float_int(res[0][0])
        else:
            return 0

    def curvepath(self, unitpos):
        """Панель отличается от прямоугольности (скос, например)"""

        sql = "SELECT curvepath FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return False

    @lru_cache(maxsize=6)
    def form(self, unitpos):
        """Форма панели
        0 - линейная
        1 - дуга по хорде
        2 - два отрезка и дуга
        """

        sql = "SELECT formtype FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return 0

    def slots_is(self, unitpos):
        """Наличие пропилов
        True - есть пропилы
        False - нет пропилов
        """
        sql = "SELECT Count(panelpos) FROM TSlots WHERE panelpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return bool(res[0][0])
        else:
            return False

    def slots_x_pc(self, unitpos):
        """Количество пропилов вдоль длины. Не учитывает напр. текстуры"""

        sql = "SELECT Count(PanelPos) FROM TSlots WHERE ROUND(BegY)=ROUND(EndY) AND PanelPos = {0}".format(
            unitpos
        )
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return 0

    def slots_y_pc(self, unitpos):
        """Количество пропилов поперёк длины. Не учитывает напр. текстуры"""

        sql = "SELECT Count(PanelPos) FROM TSlots WHERE ROUND(BegX)=ROUND(EndX) AND PanelPos = {0}".format(
            unitpos
        )
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return 0

    def slots_x_slen(self, unitpos):
        """Суммарная длина пропилов вдоль длины. Не учитывает напр. текстуры"""

        sql = "SELECT ROUND(SUM(ABS(EndX-BegX))/1000,2) FROM TSlots WHERE ROUND(BegY)=ROUND(EndY) AND PanelPos = {0}".format(
            unitpos
        )
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return 0

    def slots_y_slen(self, unitpos):
        """Суммарная длина пропилов поперёк длины. Не учитывает напр. текстуры"""

        sql = "SELECT ROUND(SUM(ABS(EndY-BegY))/1000,2) FROM TSlots WHERE ROUND(BegX)=ROUND(EndX) AND PanelPos = {0}".format(
            unitpos
        )
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return 0

    def slots_x_par(self, unitpos):
        """Параметры пропилов вдоль длины. Не учитывает напр. текстуры
        Возвращает именованный кортеж:
        'panelpos', 'beg', 'width', 'depth'
        """

        keys = ("panelpos", "beg", "width", "depth")
        sql = "SELECT PanelPos, BegY-Width/2 as Beg, Width, Depth  FROM TSlots WHERE ROUND(BegY)=ROUND(EndY) AND PanelPos = {0}".format(
            unitpos
        )
        res = self.db.rs(sql)
        lres = []
        for i in res:
            Slot = namedtuple("Slot", keys)
            i = tuple(map(utils.float_int, i))
            lres.append(Slot(*i))
        return lres

    def slots_y_par(self, unitpos):
        """Параметры пропилов поперёк длины. Не учитывает напр. текстуры
        Возвращает именованный кортеж:
        'panelpos', 'beg', 'width', 'depth'
        """

        keys = ("panelpos", "beg", "width", "depth")
        sql = (
            "SELECT PanelPos, BegX-Width/2 as Beg, Width, Depth  FROM TSlots WHERE ROUND(BegX)=ROUND(EndX) "
            "AND PanelPos = {0}".format(unitpos)
        )
        res = self.db.rs(sql)
        lres = []
        for i in res:
            Slot = namedtuple("Slot", keys)
            i = tuple(map(utils.float_int, i))
            lres.append(Slot(*i))
        return lres

    def butts_is(self, unitpos):
        """Наличие обработки торцов
        True - есть обработка
        False - нет обработки
        """
        sql = "SELECT Count(unitpos) FROM TButts WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return bool(res[0][0])
        else:
            return False

    def milling(self, up):
        """Список фрезеровок данной панели"""

        sql = "SELECT AttrString FROM TAttributes WHERE name='frezerovka' AND UnitPos = {}".format(
            up
        )
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return ""

    @lru_cache(maxsize=6)
    def priceid(self, unitpos):
        """ID элемента в номенклатурном справочнике"""

        sql = "SELECT priceid FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return 0

    def name(self, unitpos):
        """Название панели"""

        sql = "SELECT name FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return ""

    def data(self, unitpos):
        """Примечание панели"""

        sql = "SELECT data FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return ""

    def cmpos(self, unitpos):
        """Пользовательский номер элемента"""

        sql = "SELECT commonpos FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return 0

    def ppos(self, unitpos):
        """Номер родителя элемента"""

        sql = "SELECT parentpos FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return 0

    def tpos(self, unitpos):
        """Номер верхнего родителя элемента"""

        sql = "SELECT topparentpos FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return 0

    def furntype(self, unitpos):
        """Тип мебельного элемента"""

        sql = "SELECT furntype FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return "000000"

    def sumcost(self, unitpos):
        """Стоимость элемента"""

        sql = "SELECT sumcost FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if res:
            return utils.float_int(res[0][0])
        else:
            return 0

    def svw(self, unitpos):
        """Возвращает именованный кортеж
        s - Площадь кв.м
        v - объём куб.м
        w - вес кг
        """
        SVW = namedtuple("SVW", "s v w")
        id = self.priceid(unitpos)
        sql = (
            "SELECT tnpv.DValue FROM TNProperties AS tnp LEFT JOIN TNPropertyValues AS tnpv "
            "ON tnp.ID = tnpv.PropertyID WHERE (((tnp.Ident)='density') AND ((tnpv.EntityID)={}))".format(
                id
            )
        )
        res = self.db.rs(sql)
        if res:
            density = int(res[0][0])
            s = round(self.xunit(unitpos) * self.yunit(unitpos) / 1000000, 2)
            v = s * self.zunit(unitpos) / 1000
            w = v * density
            return SVW(round(s, 2), round(v, 4), round(w, 2))
        else:
            return SVW((0, 0, 0))

    def par_bent_pan(self, unitpos):
        """Возвращает именованный кортеж
        chord - хорда по концам панели
        rad - радиус панели
        h - высота панели поперёк гнутья
        ax - ось гиба OX - 1 или OY - 2
        """
        Rad = namedtuple("Rad", "rad chord h ax")
        radius = 0
        chord = 0
        form = self.form(unitpos)  # узнаём форму панели

        if form == 1:  # дуга по хорде
            sql = (
                "SELECT abs(tpm1.NumValue) AS chord, tpm2.NumValue AS chaving, tpm3.NumValue AS axis, "
                "tpm4.NumValue AS b, tpm5.NumValue AS c, tpm6.NumValue AS d, tpm7.NumValue AS e, "
                "tpm8.NumValue AS Length, tpm9.NumValue AS width FROM TParams AS tpm1, TParams AS tpm2, "
                "TParams AS tpm3, TParams AS tpm4, TParams AS tpm5, TParams AS tpm6, TParams AS tpm7, "
                "TParams AS tpm8, TParams AS tpm9 WHERE "
                "tpm1.UnitPos={0} AND tpm1.ParamName='ArcChord.Chord' "
                "AND tpm2.UnitPos={0} AND tpm2.ParamName='ArcChord.Caving' "
                "AND tpm3.UnitPos={0} AND tpm3.ParamName='BendAxis' "
                "AND tpm4.UnitPos={0} AND (tpm4.ParamName='ShavSide' AND tpm4.Hold3=3) "
                "AND tpm5.UnitPos={0} AND (tpm5.ParamName='ShavSide' AND tpm5.Hold3=1) "
                "AND tpm6.UnitPos={0} AND (tpm6.ParamName='ShavSide' AND tpm6.Hold3=0) "
                "AND tpm7.UnitPos={0} AND (tpm7.ParamName='ShavSide' AND tpm7.Hold3=2) "
                "AND tpm8.UnitPos={0} AND tpm8.ParamName='Length'"
                "AND tpm9.UnitPos={0} AND tpm9.ParamName='Width'".format(unitpos)
            )
            res = self.db.rs(sql)
            chord_0 = res[0][0]
            caving = res[0][1]
            axis = res[0][2]
            b = -res[0][3]
            c = -res[0][4]
            d = -res[0][5]
            e = -res[0][6]
            length = res[0][7]
            width = res[0][8]
            left = d
            right = e
            arc_len = width
            h = length + b + c
            if axis == 2:
                left = b
                right = c
                arc_len = length
                h = width + d + e
            min_side = min(left, right)
            max_side = max(left, right)
            l = chord_0 / 2
            radius = (l ** 2 + caving ** 2) / (2 * caving)
            alfa = arc_len * 180 / (math.pi * radius)  # угол дуги
            beta = math.radians(
                180 - (90 + (180 - alfa) / 2)
            )  # угол, что бы найти основание трапеции
            chord = chord_0 + 2 * min_side * math.cos(beta)  # основание трапеции
            if left != right:
                alfa_1 = math.radians((180 - (alfa / 2)))
                side_b = max_side - min_side
                chord = math.sqrt(
                    side_b ** 2 + chord ** 2 - 2 * side_b * chord * math.cos(alfa_1)
                )

        if form == 2:  # Два отрезка и дуга
            sql = (
                "SELECT abs(tpm1.NumValue) AS L1, tpm2.NumValue AS L2, tpm3.NumValue AS R, tpm4.NumValue AS axis, "
                "tpm5.NumValue AS b, tpm6.NumValue AS c, tpm7.NumValue AS d, tpm8.NumValue AS e, "
                "tpm9.NumValue AS Ang, tpm10.NumValue AS Length, tpm11.NumValue AS Width "
                "FROM TParams AS tpm1, TParams AS tpm2, TParams AS tpm3, TParams AS tpm4, TParams AS tpm5, "
                "TParams AS tpm6, TParams AS tpm7, TParams AS tpm8, TParams AS tpm9, "
                "TParams AS tpm10, TParams AS tpm11 WHERE "
                "tpm1.UnitPos={0} AND tpm1.ParamName='LinesArc.L1' "
                "AND tpm2.UnitPos={0} AND tpm2.ParamName='LinesArc.L2' "
                "AND tpm3.UnitPos={0} AND tpm3.ParamName='LinesArc.R' "
                "AND tpm4.UnitPos={0} AND tpm4.ParamName='BendAxis' "
                "AND tpm5.UnitPos={0} AND (tpm5.ParamName='ShavSide' AND tpm5.Hold3=3) "
                "AND tpm6.UnitPos={0} AND (tpm6.ParamName='ShavSide' AND tpm6.Hold3=1) "
                "AND tpm7.UnitPos={0} AND (tpm7.ParamName='ShavSide' AND tpm7.Hold3=0) "
                "AND tpm8.UnitPos={0} AND (tpm8.ParamName='ShavSide' AND tpm8.Hold3=2) "
                "AND tpm9.UnitPos={0} AND tpm9.ParamName='LinesArc.A' "
                "AND tpm10.UnitPos={0} AND tpm10.ParamName='Length' "
                "AND tpm11.UnitPos={0} AND tpm11.ParamName='Width'".format(unitpos)
            )

            res = self.db.rs(sql)
            len_1 = res[0][0]
            len_2 = res[0][1]
            radius = res[0][2]
            axis = res[0][3]
            b = -res[0][4]
            c = -res[0][5]
            d = -res[0][6]
            e = -res[0][7]
            ang = math.radians(res[0][8])
            length = res[0][9]
            width = res[0][10]
            left = d
            right = e
            h = length + b + c
            if axis == 2:
                left = b
                right = c
                h = width + d + e
            side_1 = len_1 + left
            side_2 = len_2 + right
            chord = math.sqrt(
                side_1 ** 2 + side_2 ** 2 - 2 * side_1 * side_2 * math.cos(ang)
            )

        chord = round(chord)
        radius = abs(round(radius))
        h = round(h, 1)
        ax = int(axis)
        return Rad(radius, chord, h, ax)

    @lru_cache(maxsize=20)
    def band_side(self, unitpos, IDLine, IDPoly=1):
        """информация о кромке торца
        Выходные данные:
        id - id кромки из номенклатуры
        name - название кромки
        width - ширина кромки
        thickpan - толщина панели
        thickband - ширина кромки
        count - кол-во кромки
        """
        Band = namedtuple(
            "Band", ["id", "name", "width", "thickpan", "thickband", "count"]
        )
        b_price_id = None  # id кромки в номенклатуре
        b_name = None  # название кромки
        b_width = None  # ширина кромки
        b_thick_pan = None  # толщина панели
        b_thick = None  # толщина кромки
        b_cnt = None

        # определим тип панели
        # PanPolyType: 1 - полилиния 2 - прямоугольная 3 - четырёхугольная
        sql = (
            "SELECT tpm.NumValue FROM TParams AS tpm WHERE (tpm.unitpos = {} AND tpm.ParamName='PanPolyType' "
            "AND tpm.HoldTable='TPanels')".format(unitpos)
        )
        res = self.db.rs(sql)
        if not res:
            return Band(0, "", 0, 0, 0, 0)
        PolyType = int(res[0][0])
        if (PolyType == 2) or (PolyType == 3):
            sql = (
                "SELECT bu.BandUnitPos, tn.ID, tn.Name, Count FROM (SELECT idp.NumValue AS IDPoly, "
                "idl.NumValue AS IDLine, idb.NumValue AS BandUnitPos FROM (TParams idp LEFT OUTER JOIN TParams idl "
                "ON idp.UnitPos=idl.UnitPos AND idp.HoldTable=idl.HoldTable AND idp.Hold1=idl.Hold1 "
                "AND idp.Hold3=idl.Hold3) LEFT OUTER JOIN TParams idb ON idp.UnitPos=idb.UnitPos "
                "AND idp.HoldTable=idb.HoldTable AND idp.Hold1=idb.Hold1 AND idp.Hold3=idb.Hold3 "
                "WHERE idp.UnitPos={0} AND idp.HoldTable='TPaths' AND idp.Hold1=1 AND idp.ParamName='IDPoly' "
                "AND idl.ParamName='IDLine' AND idb.ParamName='BandUnitPos') bu, TNNomenclature tn, TElems te "
                "WHERE bu.BandUnitPos=te.UnitPos AND te.PriceID=tn.ID AND IDPoly={1} "
                "AND IDLine={2}".format(unitpos, IDPoly, IDLine)
            )
            res = self.db.rs(sql)
            if res:
                band_id = res[0][0]
                b_price_id = res[0][1]
                b_name = res[0][2]
                b_cnt = res[0][3]
                sql = "SELECT Width FROM TBands WHERE UnitPos={}".format(band_id)
                res = self.db.rs(sql)
                if res:
                    b_thick_pan = utils.float_int(res[0][0])
                sql = (
                    "SELECT tnpv.DValue FROM TNPropertyValues AS tnpv WHERE tnpv.PropertyID=21 "
                    "AND tnpv.EntityID={}".format(b_price_id)
                )
                res = self.db.rs(sql)
                if res:
                    b_width = utils.float_int(res[0][0])
                sql = (
                    "SELECT tnpv.DValue FROM TNPropertyValues AS tnpv WHERE tnpv.PropertyID=10 "
                    "AND tnpv.EntityID={}".format(b_price_id)
                )
                res = self.db.rs(sql)
                if res:
                    b_thick = utils.float_int(res[0][0])
        return Band(b_price_id, b_name, b_width, b_thick_pan, b_thick, b_cnt)

    def band_b(self, unitpos):
        """информация о кромке торца B"""

        res = self.band_side(unitpos, 7, 1)
        return res

    def band_c(self, unitpos):
        """информация о кромке торца C"""

        res = self.band_side(unitpos, 3, 1)
        return res

    def band_d(self, unitpos):
        """информация о кромке торца D"""

        res = self.band_side(unitpos, 1, 1)
        return res

    def band_e(self, unitpos):
        """информация о кромке торца E"""

        res = self.band_side(unitpos, 5, 1)
        return res

    def band_x1(self, unitpos):
        """Кромка вдоль длины одной стороны с учётом текстуры"""

        pdir = self.dir(unitpos)
        if 45 < pdir <= 135:
            return self.band_b(unitpos)
        elif 225 < pdir <= 315:
            return self.band_c(unitpos)
        else:
            return self.band_e(unitpos)

    def band_x2(self, unitpos):
        """Кромка вдоль длины одной стороны с учётом текстуры"""

        pdir = self.dir(unitpos)
        if 45 < pdir <= 135:
            return self.band_c(unitpos)
        elif 225 < pdir <= 315:
            return self.band_b(unitpos)
        else:
            return self.band_d(unitpos)

    def band_y1(self, unitpos):
        """Кромка вдоль длины одной стороны с учётом текстуры"""

        pdir = self.dir(unitpos)
        if 45 < pdir <= 135:
            return self.band_e(unitpos)
        elif 225 < pdir <= 315:
            return self.band_d(unitpos)
        else:
            return self.band_b(unitpos)

    def band_y2(self, unitpos):
        """Кромка вдоль длины одной стороны с учётом текстуры"""

        pdir = self.dir(unitpos)
        if 45 < pdir <= 135:
            return self.band_d(unitpos)
        elif 225 < pdir <= 315:
            return self.band_e(unitpos)
        else:
            return self.band_c(unitpos)

    def total_bands_to_panels(self, pans):
        """Сумарное количество кромок, принадлежащих списку деталей.
        Вывод: именованный кортеж priceid, length
        """

        sql = (
            "SELECT te.PriceID, Round(Sum(tb.Length * te.Count)/1000, 2) AS Length "
            "FROM TBands AS tb RIGHT JOIN TElems AS te "
            "ON tb.UnitPos = te.UnitPos WHERE te.ParentPos in {} "
            "GROUP BY te.PriceID, te.FurnType HAVING te.FurnType='050000'".format(
                tuple(pans)
            )
        )
        res = self.db.rs(sql)
        lres = []
        if res:
            for i in res:
                TBP = namedtuple("TBP", "priceid length")
                lres.append(TBP(*i))
        return lres

    def decorates(self, up, map=5):
        """Определяет отделку панели выбранной карты map:
        1 - сторона E (Y+)
        2 - сторона D (Y-)
        3 - сторона C (X+)
        4 - сторона B (X-)
        5 - пласть A (Z+)
        6 - пласть F (Z-)
        7 - угол 1
        8 - угол 2
        9 - угол 3
        10 - угол 4
        11 - дополнение 1
        12 - дополнение 2
        -1 - все стороны
        -2 - все торцы"""

        sql = (
            "SELECT td.materialid, td.typename, tnn.name FROM TDecorates AS td LEFT JOIN TNNomenclature AS tnn "
            "ON td.materialid=tnn.id WHERE Unitpos={} AND map={}".format(up, map)
        )
        res = self.db.rs(sql)
        abr = {
            "шпон": "шп.",
            "эмаль": "эм.",
            "пластик": "пл.",
            "плёнка (пвх)": "ПВХ ",
            "патина": "пт.",
            "морилка": "мор.",
            "лак": "",
        }
        rep = ["шп.", "шпон", "Шпон", "эм.", "Эм.", "эмаль", "Эмаль"]
        dec = []
        for i in res:
            dec.append(i[2])
        return " ".join(dec).strip()

    def cnt_drill_pans(self, tpp=None, hingoff=False):
        """Получить кол-во просверленных деталей. Сверловка с двух сторон считается, как две детали
        hingoff = True - Считать вместе с присадкой под петли
        """
        filtrtpp = "AND TElems.TopParentPos={0}".format(tpp) if tpp is not None else ""
        # Выбор сверловки всего, кроме стороны F
        sql_exf = (
            "SELECT DISTINCT th.UnitPos FROM THoles AS th LEFT JOIN TElems ON th.UnitPos = TElems.UnitPos "
            "WHERE (th.MatrA33<=0 AND (th.UnitPos In (SELECT UnitPos FROM TPanels)) {0})".format(
                filtrtpp
            )
        )
        # Выбор отвертий в пласти F
        sql_f = (
            "SELECT DISTINCT th.UnitPos FROM THoles AS th LEFT JOIN TElems ON th.UnitPos = TElems.UnitPos "
            "WHERE (Abs(th.MatrA13)<0.001 AND Abs(th.MatrA23)<0.001 AND th.MatrA33>0 "
            "AND th.UnitPos In (SELECT UnitPos FROM TPanels) {0})".format(filtrtpp)
        )

        sql_un = " UNION ".join([sql_exf, sql_f])
        sql = "SELECT Count(*) FROM ({})".format(sql_un)
        res = self.db.rs(sql)
        if res:
            if hingoff:
                return res[0][0] - self.cnt_pan_hings(tpp=tpp)
            else:
                return res[0][0]

    def cnt_holes_pan(self, unitpos, hingoff=False):
        """Получить кол-во отверстий конкретной панели
        hingoff - Считать вместе с присадкой под петли или исключить петли
        """
        sql = "SELECT Count(*) FROM THoles WHERE THoles.UnitPos={0}".format(unitpos)
        res = self.db.rs(sql)
        if not res:
            return 0
        if hingoff:
            return res[0][0] - self.cnt_holes_hings(unitpos=unitpos)
        else:
            return res[0][0]

    def cnt_holes_pan_diam(self, unitpos, d):
        """Получить кол-во отверстий определённого диаметра конкретной панели"""

        sql = "SELECT Count(*) FROM THoles WHERE THoles.UnitPos={0} AND THoles.Diameter={1}".format(
            unitpos, d
        )
        res = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]

    def cnt_holes_hings(self, unitpos=None, tpp=None):
        """Получить кол-во отверстий под чашку пети у всех панелей,
        при unitpos=None или конкретной панели, при unitpos=UnitPos панели.
        tpp=TopParentPos - id родительского объекта
        """
        if unitpos and tpp is None:
            filter_up = "AND th.UnitPos = {0} ".format(unitpos)
        else:
            filter_up = ""
        if tpp:
            filter_tpp = "AND te.TopParentPos = {0}".format(tpp)
        else:
            filter_tpp = ""
        sql = (
            "SELECT Count(th.UnitPos) FROM THoles AS th LEFT JOIN TElems AS te ON th.HolderPos = te.UnitPos "
            "WHERE te.FurnType Like '0406%' AND th.Diameter > 20 {0}".format(
                filter_up + filter_tpp
            )
        )
        res = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]

    def cnt_hings_x(self, unitpos):
        """Количество петель вдоль панели"""

        sql = (
            "SELECT Count(th.MatrA24) FROM THoles AS th LEFT JOIN TElems AS te ON th.HolderPos = te.UnitPos "
            "WHERE te.FurnType Like '0406%' AND th.Diameter > 20 AND th.UnitPos = {0} "
            "GROUP BY th.MatrA24 HAVING Count(1)>1".format(unitpos)
        )
        res = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]

    def cnt_hings_y(self, unitpos):
        """Количество петель поперёк панели"""

        sql = (
            "SELECT Count(th.MatrA14) FROM THoles AS th LEFT JOIN TElems AS te ON th.HolderPos = te.UnitPos "
            "WHERE te.FurnType Like '0406%' AND th.Diameter > 20 AND th.UnitPos = {0} "
            "GROUP BY th.MatrA14 HAVING Count(1)>1".format(unitpos)
        )
        res = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]

    @lru_cache(maxsize=6)
    def cnt_pan_hings(self, tpp=None):
        """Получить кол-во деталей с присадкой только под петли (без других отв.)
        tpp=TopParentPos - id родительского объекта
        """
        if tpp:
            filter_tpp = "AND te.TopParentPos = {0}".format(tpp)
        else:
            filter_tpp = ""
        sql = (
            "SELECT Count(*) FROM (SELECT DISTINCT th.UnitPos FROM THoles AS th LEFT JOIN TElems AS te "
            "ON th.HolderPos = te.UnitPos WHERE (te.FurnType) Like '0406%' {0} AND th.UnitPos not in (SELECT "
            "th.UnitPos FROM THoles AS th LEFT JOIN TElems AS te ON th.HolderPos = te.UnitPos "
            "WHERE (te.FurnType) not Like '0406%' {0}))".format(filter_tpp)
        )
        res = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]

    def cnt_chamfer_pan(self, unitpos):
        """Определение кол-ва скосов углов на панеле"""

        sql = (
            "SELECT Count(*) FROM TParams WHERE ParamName = 'Cuttype' AND NumValue = 1 "
            "AND UnitPos = {0}".format(unitpos)
        )
        res = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]

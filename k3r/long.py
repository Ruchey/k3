# -*- coding: utf-8 -*-


from collections import namedtuple
from functools import lru_cache
from . import utils

__author__ = 'Виноградов А.Г. г.Белгород  август 2015'


class Long:
    """Класс работы с длиномерами"""

    def __init__(self, db):
        self.db = db

    def form(self, up):
        """Возвращает форму длиномера
        Входные данные: up - UnitPos в таблице TLongs
        Выходные данные:
        0 - линейная
        1 - дуга по хорде
        2 - два отрезка и дуга
        """
        sql = "SELECT tl.LongTable, te.UnitPos FROM TLongs AS tl LEFT JOIN TElems AS te " \
              "ON tl.UnitPos = te.ParentPos WHERE tl.UnitPos={}".format(up)
        res = self.db.rs(sql)
        table = res[0][0]
        unit_pos = res[0][1] if res[0][1] else up
        sql = "SELECT FormType FROM {0} AS tb WHERE tb.UnitPos={1}".format(table, unit_pos)
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return 0

    @lru_cache(maxsize=6)
    def long_list(self, lt=None, tpp=None):
        """Возвращает список длиномеров именованым кортежем
        Входные данные:
        lt - LongType тип длиномера
        tpp - TopParentPos хозяин
        Вывод: 'type', 'priceid', 'length', 'width', 'height', 'cnt', 'form
        Типы длиномеров:
            0	Столешница
            1	Карниз
            2	Стеновая панель
            3	Водоотбойник
            4	Профиль карниза
            5	Цоколь
            6	Нижний профиль
            7	Балюстрада
        """
        keys = ('type', 'priceid', 'goodsid', 'length', 'width', 'height', 'cnt', 'form')
        gr_keys = ('type', 'priceid', 'goodsid', 'length', 'width', 'height', 'form')
        filter_lt = "WHERE LongType={}".format(lt) if not lt is None else ""
        pref = " AND" if lt else "WHERE"
        filter_tpp = "{} te.TopParentPos={}".format(pref, tpp) if tpp else ""
        sql = "SELECT tl.UnitPos, tl.LongType AS lt, te.PriceID, tl.LongGoodsID, " \
              "te.XUnit, te.YUnit, te.ZUnit, te.Count FROM TLongs AS tl INNER JOIN TElems AS te " \
              "ON tl.UnitPos = te.UnitPos {} ORDER BY tl.LongType".format(filter_lt + filter_tpp)
        res = self.db.rs(sql)
        d_res = []
        for i in res:
            long = namedtuple('Long', keys)
            i += (self.form(i[0]),)
            i_lst = list(i)
            i_lst.pop(0)
            d_res.append(long(*i_lst))
        gr_lst = utils.group_by_keys(d_res, gr_keys, 'cnt')
        return gr_lst

    @lru_cache(maxsize=6)
    def total(self, lt=None, tpp=None):
        """Суммарное колличество длиномеров согласно единицам измерения материалов
        LongType, LongMatID, Length, LongGoodsID
        Входные данные:
        lt - LongType тип длиномера
        tpp - TopParentPos хозяин
        Вывод: "type", "priceid", "form", 'goodsid', "quantity"
        """
        keys = ("type", "priceid", "form", 'goodsid', "quantity")
        longs = self.long_list(lt, tpp)
        total = []
        gr_longs = {}
        for i in longs:
            key = (i.type, i.priceid, i.form, i.goodsid)
            if not key in list(gr_longs.keys()):
                gr_longs[key] = []
            gr_longs[key].append((i.length, i.cnt))
        for i in gr_longs:
            form = i[2]
            if form == 0:
                quantity = sum(length*cnt for length, cnt in gr_longs[i])
            else:
                quantity = sum(cnt for length, cnt in gr_longs[i])
            total_long = namedtuple("TLong", keys)
            total.append(total_long(*(i + (quantity,))))
        return total

    def long(self):
        pass

# -*- coding: utf-8 -*-
__author__ = 'Виноградов А.Г. г.Белгород  август 2015'

from collections import namedtuple


class Profile:
    """Класс работы с профилями"""

    def __init__(self, db):
        self.db = db

    def profiles(self, tpp=None):
        """Возвращает список именованных кортежей профилей.
        Не включает профиля длиномеров
        unitpos, length, colorid, formtype
        """
        keys = ('unitpos', 'length', 'colorid', 'formtype')
        filter_tpp = " AND te.TopParentPos={}".format(tpp)
        if tpp is None:
            filter_tpp = ""
        sql = "SELECT tpf.UnitPos, tpf.Length, tpf.ColorID, tpf.FormType FROM TProfiles AS tpf " \
              "INNER JOIN TElems AS te ON tpf.UnitPos = te.UnitPos " \
              "WHERE NOT EXISTS(SELECT te.UnitPos FROM TElems AS te, TLongs AS tl " \
              "WHERE te.ParentPos=tl.UnitPos AND te.UnitPos=tpf.UnitPos){}".format(filter_tpp)
        res = self.db.rs(sql)
        l_res = []
        if res:
            for i in res:
                Prof = namedtuple('Prof', keys)
                l_res.append(Prof(*i))
        return l_res

    def total(self, tpp=None):
        """Возвращает общее количество профилей
        priceid, length, colorid, formtype
        thick - толщина пилы прибавляемая к каждому отрезку
        """
        keys = ('priceid', 'length', 'colorid', 'formtype')
        filter_tpp = " AND te.TopParentPos={}".format(tpp)
        if tpp is None:
            filter_tpp = ""
        thick = 4
        sql = "SELECT te.PriceID, Sum(([tpf].[Length]+{1})*[te].[Count]) AS Length, tpf.ColorID " \
              "FROM TProfiles AS tpf INNER JOIN TElems AS te ON tpf.UnitPos = te.UnitPos " \
              "WHERE (((Exists (SELECT te.UnitPos FROM TElems AS te, TLongs AS tl " \
              "WHERE te.ParentPos=tl.UnitPos AND te.UnitPos=tpf.UnitPos))=False)){0} " \
              "GROUP BY te.PriceID, tpf.ColorID".format(filter_tpp, thick)
        res = self.db.rs(sql)
        l_res = []
        if res:
            for i in res:
                Prof = namedtuple('Prof', keys)
                l_res.append(Prof(*i))
        return l_res

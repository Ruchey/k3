# -*- coding: utf-8 -*-
__author__ = 'Виноградов А.Г. г.Белгород  август 2015'


class Profile:
    '''Класс работы с профилями'''
    def __init__(self, db):
        try:
            self.db = db
        except:
            return None

    def pflist(self, tpp=None):
        '''Возвращает список профилей. Не включает профиля длиномеров
            UnitPos, Length, ColorID, FormType'''
        filtrtpp = " AND te.TopParentPos={}".format(tpp)
        if tpp is None:
            filtrtpp = ""
        sql = "SELECT tpf.UnitPos, tpf.Length, tpf.ColorID, tpf.FormType FROM TProfiles AS tpf " \
              "INNER JOIN TElems AS te ON tpf.UnitPos = te.UnitPos " \
              "WHERE NOT EXISTS(SELECT te.UnitPos FROM TElems AS te, TLongs AS tl " \
              "WHERE te.ParentPos=tl.UnitPos AND te.UnitPos=tpf.UnitPos){}".format(filtrtpp)
        res = self.db.recordset(sql)
        return res

    def sumcount(self, tpp=None):
        '''Возвращает общее количество профилей
            PriceID, Length, ColorID, FormType
            thick - толщина пилы прибавляемая к каждому отрезку'''
        filtrtpp = " AND te.TopParentPos={}".format(tpp)
        if tpp is None:
            filtrtpp = ""
        thick = 4
        sql = "SELECT te.PriceID, Sum(([tpf].[Length]+{1})*[te].[Count]) AS Length, tpf.ColorID " \
              "FROM TProfiles AS tpf INNER JOIN TElems AS te ON tpf.UnitPos = te.UnitPos " \
              "WHERE (((Exists (SELECT te.UnitPos FROM TElems AS te, TLongs AS tl " \
              "WHERE te.ParentPos=tl.UnitPos AND te.UnitPos=tpf.UnitPos))=False)){0} " \
              "GROUP BY te.PriceID, tpf.ColorID".format(filtrtpp, thick)
        res = self.db.recordset(sql)
        return res
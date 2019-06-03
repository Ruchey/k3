# -*- coding: utf-8 -*-
__author__ = 'Виноградов А.Г. г.Белгород  август 2015'


class Longs:
    '''Класс работы с длиномерами'''
    def __init__(self, db):
        try:
            self.db = db
        except:
            return None

    def llist(self, lt=None, tpp=None):
        '''Возвращает список длиномеров
            UnitPos, LongType, LongTable, LongMatID, LongGoodsID
            Входные данные:
            lt - LongType тип длиномера
            tpp - TopParentPos хозяин'''
        filtrlt = ("WHERE LongType={}".format(lt) if lt else "")
        pref = (" AND" if lt else "WHERE")
        filtrtpp = ("{} te.TopParentPos={}".format(pref, tpp) if tpp else "")
        sql = "SELECT tl.UnitPos, tl.LongType, tl.LongTable, tl.LongMatID, tl.LongGoodsID FROM TLongs AS tl " \
              "INNER JOIN TElems AS te ON tl.UnitPos = te.UnitPos {} ORDER BY tl.LongType".format(filtrlt+filtrtpp)
        res = self.db.recordset(sql)
        
        return res

    def sumcount(self, lt=None, tpp=None):
        '''Суммарное колличество длиномеров согласно
            единицам измерения материалов
             LongType, LongMatID, Length, LongGoodsID
            Входные данные:
            lt - LongType тип длиномера
            tpp - TopParentPos хозяин'''
        longs = Longs.llist(self,lt, tpp)
        nlst = {}
        sc = []
        for i in longs:
            if not i[-4:] in list(nlst.keys()):
                nlst[i[-4:]] = []
            nlst[i[-4:]].append(i[0])
        for i in nlst.items():
            sqlPan = "SELECT Sum([tp].[Length]*[tp].[Width]/10^6*[te].[Count]) AS Cnt FROM TElems AS te " \
                     "INNER JOIN TPanels AS tp ON te.UnitPos = tp.UnitPos WHERE te.ParentPos in {}".format(tuple(i[1]))

            sqlPf = "SELECT Sum([tpf].[Length]/10^3*[te].[Count]) AS Cnt " \
                    "FROM TElems AS te INNER JOIN TProfiles AS tpf ON te.UnitPos = tpf.UnitPos " \
                    "WHERE te.ParentPos in {}".format(tuple(i[1]))

            sqlBl = "SELECT Sum([tbl].[Length]/10^3*[te].[Count]) AS Cnt " \
                    "FROM TElems AS te INNER JOIN TBalusters AS tbl ON te.UnitPos = tbl.UnitPos " \
                    "WHERE tbl.UnitPos in {}".format(tuple(i[1]))

            dsql = {'TPanels':sqlPan, 'TProfiles':sqlPf, 'TBalusters':sqlBl}
            sql = dsql[i[0][1]]
            res = self.db.recordset(sql)[0][0]
            sc.append([i[0][0], i[0][2], res, i[0][3]])
        return sc
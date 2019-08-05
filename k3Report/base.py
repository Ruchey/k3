# -*- coding: utf-8 -*-

from . import utils
from collections import namedtuple


__author__ = 'Виноградов А.Г. г.Белгород  август 2015'


class Base:
    """Класс работы с таблицами базы выгрузки"""
    def __init__(self, db):
        try:
            self.db = db
        except:
            return None

    def torderinfo(self):
        """Выводит именованный кортеж данных из таблицы TOrderInfo"""
        
        keys = ('orderid', 'ordername', 'ordernumber', 'customer', 'address', \
                'telephonenumber', 'orderdata', 'orderexpirationdata', 'firm', \
                'salon', 'acceptor', 'executor', 'additionalinfo', 'toworking')
        TOrder = namedtuple('TOrder', keys)
        sql = "SELECT * FROM TOrderInfo"
        res = self.db.rs(sql)
        if res:
            return TOrder(*res[0])
        else:
            return res

    def tobjects(self, up=None):
        """Данные из таблицы объектов TObjects"""
        
        keys = ('UnitPos', 'PlaceType', 'Article', 'ExArticle', 'IsStandart', \
                'Cataloque', 'Library', 'LibName', 'LibCaption', 'ProtoID', 'BasePrice')
        filtrup = ('WHERE UnitPos = {}'.format(up) if up else '')
        sql = "SELECT * FROM TObjects {}".format(filtrup)
        res = self.db.rs(sql)
        dres = []
        for i in res:
            dres.append(dict(zip(keys,i)))
        return dres

    def telems(self, UnitPos):
        
        keys = ('parentpos', 'topparentpos', 'detailpos', 'commonpos', \
                'levelpos', 'furntype', 'furnkind', 'name', 'priceid', \
                'goodsid', 'sumcost', 'xunit', 'yunit', 'zunit', 'count', \
                'data', 'hashcode')
        TElems = namedtuple('TElems', keys)
        sql = "SELECT ParentPos, TopParentPos, DetailPos, CommonPos, LevelPos, "\
              "FurnType, FurnKind, [Name], PriceID, " \
              "GoodsID, SumCost, XUnit, YUnit, ZUnit, [Count], [Data], " \
              "HashCode FROM TElems WHERE UnitPos={}".format(UnitPos)
        res = self.db.rs(sql)
        if res:
            return TElems(*res[0])
        else:
            return res

    def tdrawings(self, up=None):
        """Все данные из таблицы TDrawings"""
        filtrup = ('WHERE UnitPos = {}'.format(up) if up else '')
        sql = "SELECT * FROM TDrawings {}".format(filtrup)
        res = self.db.rs(sql)
        keys = ('UnitPos', 'DrawingName', 'DrawingDescr', 'SizeX', 'SizeY')
        dres = []
        for i in res:
            dres.append(dict(zip(keys,i)))
        return dres

    def tattributes(self, up):
        """Таблица атрибутов"""
        sql ="SELECT atr.Name, Switch(AttrType=1,AttrString,AttrType=2,AttrReal,AttrType=3,AttrText,AttrType=4,AttrReal) AS val " \
             "FROM TAttributes AS atr WHERE atr.UnitPos={}".format(up)
        res = dict(self.db.rs(sql))
        return res

    def tngoods(self, id=None):
        """Данные из таблицы TNGoods"""
        filtrid = ('WHERE ID = {}'.format(id) if id else '')
        sql = "SELECT * FROM TNGoods {}".format(filtrid)
        res = self.db.rs(sql)
        keys = ('ID', 'Name', 'GroupID', 'GroupName', 'FurnType', 'ParentID', 'GLevel')
        dres = []
        for i in res:
            dres.append(dict(zip(keys,i)))
        return dres

    def tnnomenclature(self, id):
        """Данные таблицы номенклатуры"""
        sql = "SELECT * FROM TNNomenclature WHERE ID = {0}".format(id)
        keys = ('ID', 'Name', 'MatTypeID', 'MatTypeName', 'GroupID', 'GroupName', 'KindID', \
                'KindName', 'Article', 'UnitsID', 'UnitsName', 'Price', 'ParentID', 'GLevel')
        val = self.db.rs(sql)
        res = dict(zip(keys,val[0]))
        return res
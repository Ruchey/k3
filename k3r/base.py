# -*- coding: utf-8 -*-

from . import utils
from collections import namedtuple


__author__ = "Виноградов А.Г. г.Белгород  август 2015"


class Base:
    """Класс работы с таблицами базы выгрузки"""

    def __init__(self, db):
        self.db = db

    def torderinfo(self):
        """
        Выводит именованный кортеж данных из таблицы TOrderInfo
        """
        keys = (
            "orderid",
            "ordername",
            "ordernumber",
            "customer",
            "address",
            "telephonenumber",
            "orderdata",
            "orderexpirationdata",
            "firm",
            "salon",
            "acceptor",
            "executor",
            "additionalinfo",
            "toworking",
        )
        TOrder = namedtuple("TOrder", keys)
        sql = "SELECT * FROM TOrderInfo"
        res = self.db.rs(sql)
        if res:
            return TOrder(*res[0])
        else:
            return res

    def tobjects(self, up=None):
        """
        Данные из таблицы объектов TObjects
        """
        keys = (
            "unitpos",
            "placetype",
            "article",
            "exarticle",
            "isstandart",
            "cataloque",
            "library",
            "libname",
            "libcaption",
            "protoid",
            "baseprice",
        )
        TObjects = namedtuple("TObjects", keys)
        objects = []
        if up:
            sql = "SELECT * FROM TObjects WHERE UnitPos = {}".format(up)
        else:
            sql = "SELECT * FROM TObjects"
        res = self.db.rs(sql)
        if res and up:
            return TObjects(*res[0])
        elif res:
            for obj in res:
                objects.append(TObjects(*obj))
            return objects
        else:
            return res

    def telems(self, unitpos):

        keys = (
            "parentpos",
            "topparentpos",
            "detailpos",
            "commonpos",
            "levelpos",
            "furntype",
            "furnkind",
            "name",
            "priceid",
            "goodsid",
            "sumcost",
            "xunit",
            "yunit",
            "zunit",
            "count",
            "data",
            "hashcode",
        )
        TElems = namedtuple("TElems", keys)
        sql = (
            "SELECT ParentPos, TopParentPos, DetailPos, CommonPos, LevelPos, "
            "FurnType, FurnKind, [Name], PriceID, "
            "GoodsID, SumCost, XUnit, YUnit, ZUnit, [Count], [Data], "
            "HashCode FROM TElems WHERE UnitPos={}".format(unitpos)
        )
        res = self.db.rs(sql)
        if res:
            return TElems(*res[0])
        else:
            return res

    def tdrawings(self, unitpos=None):
        """
        Все данные из таблицы TDrawings
        """
        keys = ("unitpos", "drawingname", "drawingdescr", "sizex", "sizey")
        TDrawings = namedtuple("TDrawings", keys)
        filtrup = "WHERE UnitPos = {}".format(unitpos) if unitpos else ""
        sql = "SELECT * FROM TDrawings {}".format(filtrup)
        res = self.db.rs(sql)
        if res:
            return TDrawings(*res[0])
        else:
            return res

    def tattributes(self, unitpos):
        """
        Таблица атрибутов
        """
        sql = (
            "SELECT atr.Name, Switch(AttrType=1, AttrString, AttrType=2, AttrReal, AttrType=3,AttrText, "
            "AttrType=4,AttrReal) AS val FROM TAttributes AS atr WHERE atr.UnitPos={}".format(
                unitpos
            )
        )
        res = dict(self.db.rs(sql))
        return res

    def tngoods(self, id_=None):
        """
        Данные из таблицы TNGoods
        """
        keys = ("id", "name", "groupid", "groupname", "furntype", "parentid", "glevel")
        TNGoods = namedtuple("TNGoods", keys)
        filtrid = "WHERE ID = {}".format(id_) if id_ else ""
        sql = "SELECT * FROM TNGoods {}".format(filtrid)
        res = self.db.rs(sql)
        if res:
            return TNGoods(*res[0])
        else:
            return res

    def tnnomenclature(self, id_):
        """
        Данные таблицы номенклатуры
        """
        keys = (
            "id",
            "name",
            "mattypeid",
            "mattypename",
            "groupid",
            "groupname",
            "kindid",
            "kindname",
            "article",
            "unitsid",
            "unitsname",
            "price",
            "parentid",
            "glevel",
        )
        TNNomenclature = namedtuple("TNNomenclature", keys)
        sql = "SELECT * FROM TNNomenclature WHERE ID = {0}".format(id_)
        res = self.db.rs(sql)
        if res:
            return TNNomenclature(*res[0])
        else:
            return res

    def get_anc_furntype(self, up, ft):
        """Возвращает родителя элемента с UnitPos = up по правилу LIKE ft*"""

        sql = (
            "SELECT te.UnitPos, te.ParentPos, te.TopParentPos, te.FurnType "
            "FROM TElems AS te LEFT JOIN TElems AS te1 ON te.TopParentPos = te1.TopParentPos "
            "WHERE te1.UnitPos={}".format(up)
        )
        res = self.db.rs(sql)
        tree = utils.get_tree_parents(up, res)
        for i in tree:
            if i[3][:2] == ft:
                return i
        return []

    def get_child_furntype(self, up, ft, top=1):
        """Возвращает потомка элемента с UnitPos = up по правилу LIKE ft*
        Если UnitPos не задан, т.е. None, вернёт всех по FurnType LIKE ft*
        Входные данные:
            up -- int UnitPos
            ft - str FurnType
            top -- bool, если True, то поиск по верхнему родителю TopParentPos,
                    иначе поиск по ParentPos
        Вывод:
            [(UnitPos, FurnType), (UnitPos, FurnType)]
            [(598, '010300'), (599, '010100')]
        """
        if top:
            parent = "te.TopParentPos={0}".format(up)
        else:
            parent = "te.ParentPos={0}".format(up)
        sql = (
            "SELECT te.UnitPos, te.FurnType FROM TElems AS te "
            "WHERE {0} AND te.FurnType Like '{1}%'".format(parent, ft)
        )
        if not up:
            sql = (
                "SELECT te.UnitPos, te.FurnType FROM TElems AS te "
                "WHERE te.FurnType Like '{0}%'".format(ft)
            )
        res = self.db.rs(sql)
        return res

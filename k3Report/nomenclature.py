# -*- coding: utf-8 -*-

from . import utils
from collections import namedtuple


__author__ = 'Виноградов А.Г. г.Белгород  август 2015'


class Nomenclature:
    """Класс работы с номенклатурой"""
    
    def __init__(self, db):

        self.db = db

    def accbyuid(self, uid=None, tpp=None):
        """Выводит список аксессуаров
            uid = id единицы измерения
            tpp = TopParentPos ID верхнего хозяина аксессуара
            Вывод: id, name, article, unitsname, count, price
        """
        filtruid = "AND tnn.UnitsID={}".format(uid) if uid else ""
        filtrtpp = "AND te.TopParentPos={}".format(tpp) if tpp else ""
        where = " ".join(["tnn.UnitsID<>11", filtruid, filtrtpp])
        keys = ('id', 'name', 'article', 'unitsname', 'count', 'price')

        sql = "SELECT ID, [Name], Article, UnitsName, Sum(Cnt/IIf(UnitsID=10,accCnt,1)), Price FROM " \
              "(SELECT tnn.ID, tnn.Name, tnn.Article, tnn.UnitsID, tnn.UnitsName, (te.Count) AS Cnt, tnn.Price, " \
              "(Select (Max(ta1.AccType)-Min(ta1.AccType))+1 from TAccessories AS ta1 where ta1.AccMatID=ta.AccMatID) AS accCnt " \
              "FROM (TAccessories AS ta LEFT JOIN TNNomenclature AS tnn ON ta.AccMatID = tnn.ID) LEFT JOIN TElems AS te " \
              "ON ta.UnitPos = te.UnitPos WHERE {0}) GROUP BY ID, [Name], Article, UnitsName, Price " \
              "ORDER BY [Name]".format(where)
        res = self.db.rs(sql)
        lres = []
        for i in res:
            Ac = namedtuple('Ac', keys)
            lres.append(Ac(*i))
        return lres

    def acclong(self, tpp=None):
        """Список погонажных комплектующих, таких как сетки
           tpp = TopParentPos ID верхнего хозяина аксессуара
           ID, Название, артикль, ед.изм., длина, кол-во, цена
        """
        filtrtpp = ('WHERE te.TopParentPos={}'.format(tpp) if tpp else '')
        keys = ('id', 'name', 'article', 'unitsname', 'length', 'count', 'price')
        sql = "SELECT tnn.ID, tnn.Name, tnn.Article, tnn.UnitsName, te.XUnit/1000, te.Count, tnn.Price FROM TElems AS te " \
              "INNER JOIN TNNomenclature AS tnn ON te.PriceID = tnn.ID WHERE te.FurnType Like '07%' {0} ORDER BY te.Name".format(filtrtpp)
        res = self.db.rs(sql)
        lres = []
        for i in res:
            Ac = namedtuple('Ac', keys)
            lres.append(Ac(*i))
        return lres

    def matbyuid(self, uid, tpp=None):
        """Выводит список ID материалов
            uid = id единицы измерения
            tpp = TopParentPos ID верхнего хозяина материала
        """
        where = "{0} AND te.TopParentPos={1}".format(uid, tpp) if tpp else "{}".format(uid)
        sql = "SELECT tnn.ID FROM TElems AS te INNER JOIN TNNomenclature AS tnn ON te.PriceID = tnn.ID " \
              "WHERE tnn.UnitsID={} GROUP BY tnn.ID".format(where)
        
        res = self.db.rs(sql)
        if res:
            id = []
            for i in res:
                id.append(i[0])
            return id
        else:
            return res

    def properties(self, id):
        """Возвращает именованный кортеж свойств номенклатурной единицы
        properties.name
        properties.article
        properties.unitsid
        properties.unitsname
        properties.price
        properties.mattypeid
        properties.
        """
        
        keys = ('id', 'name', 'article', 'unitsid', 'unitsname', 'price', 'mattypeid')
        frm = "TNProperties AS tnp INNER JOIN TNPropertyValues AS tnpv ON tnp.ID = tnpv.PropertyID"
        
        sql1 = "SELECT LCase(tnp.Ident), tnpv.DValue FROM {0} WHERE (tnp.TypeID in (1,7) AND tnpv.EntityID={1});".format(frm, id)
        sql2 = "SELECT LCase(tnp.Ident), tnpv.IValue FROM {0} WHERE (tnp.TypeID in (3,6,11,17,18) AND tnpv.EntityID={1});".format(frm, id)
        sql3 = "SELECT LCase(tnp.Ident), tnpv.SValue FROM {0} WHERE (tnp.TypeID in (5,12,13,14,15,16) AND tnpv.EntityID={1});".format(frm, id)
        sql4 = "SELECT ID, Name, Article, UnitsID, UnitsName, Price, MatTypeID FROM TNNomenclature AS tnn WHERE tnn.ID={0}".format(id)
        
        res1 = self.db.rs(sql1)
        res2 = self.db.rs(sql2)
        res3 = self.db.rs(sql3)
        res4 = self.db.rs(sql4)
        if res4:
            res4 = list(zip(keys,self.db.rs(sql4)[0]))
        res = utils.normkeyprop(res1 + res2 + res3 + res4)
        if res:
            Prop = namedtuple('Prop', [x[0] for x in res])
            return Prop(**dict(res))
        else:
            return None

    def property_name(self, id, prop):
        """Получить значение именованного сво-ва
           id - номер материала
           prop - имя свойства
        """
        keys = ('ID', 'Name', 'MatTypeID', 'MatTypeName', 'GroupID', 'GroupName', 'KindID',
                'KindName', 'Article', 'UnitsID', 'UnitsName', 'Price', 'ParentID', 'GLevel')
        sql1 = "SELECT Switch([tnp].[TypeID]=1,[tnpv].[DValue],[tnp].[TypeID]=3,[tnpv].[IValue]," \
              "[tnp].[TypeID]=5,[tnpv].[SValue],[tnp].[TypeID]=6,[tnpv].[IValue],[tnp].[TypeID]=7,[tnpv].[DValue]," \
              "[tnp].[TypeID]=11,[tnpv].[IValue],[tnp].[TypeID]=12,[tnpv].[SValue],[tnp].[TypeID]=13,[tnpv].[SValue]," \
              "[tnp].[TypeID]=14,[tnpv].[SValue],[tnp].[TypeID]=15,[tnpv].[SValue],[tnp].[TypeID]=16,[tnpv].[SValue]," \
              "[tnp].[TypeID]=17,[tnpv].[IValue],[tnp].[TypeID]=18,[tnpv].[IValue]) AS val " \
              "FROM TNProperties AS tnp INNER JOIN TNPropertyValues AS tnpv ON tnp.ID = tnpv.PropertyID " \
              "WHERE tnpv.EntityID={0} AND tnp.Ident='{1}'".format(id, prop)
        sql2 = "SELECT tnn.{1} FROM TNNomenclature AS tnn WHERE tnn.ID={0}".format(id, prop)
        if prop in keys:
            sql = sql2
        else:
            sql = sql1
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return None

    def sqm(self, id, tpp=None):
        "Определяем кол-во площадного материала"
        
        where = "{0} AND te.TopParentPos={1}".format(id, tpp) if tpp else "{}".format(id)
        sql = "SELECT Round(Sum((XUnit*YUnit*Count)/10^6), 2) AS sqr FROM TElems AS te WHERE te.PriceID={}".format(where)
        res = self.db.rs(sql)
        if res:
            return res[0][0]
        else:
            return 0

    def matcount(self, id, tpp=None):
        "Выводит кол-во материала"
        
        cnt = 0
        # Проверим, является ли материал аксессуаром
        sql = "SELECT Count(AccMatID) FROM TAccessories WHERE TAccessories.AccMatID={}".format(id)
        res = self.db.rs(sql)
        if res:
            if res[0][0]>0:
                cnt = self.accbyuid(id, tpp)
            else:
                unit = self.property_name(id, 'UnitsID')
                if unit == 2: # Квадратные метры
                    cnt = self.sqm(id, tpp)
        return cnt

    def bands(self, add=0, tpp=None):
        """Информация по кромке.
           Входные данные:
           add - добавочная длина кромки в мм на торец для отходов
           tpp - ID хозяина кромки
           Выходные данные:
           id - ID материала
           length - длина
           thick - толщина торца
        """
        filtrtpp = ("WHERE te.TopParentPos={}".format(tpp) if tpp else '')
        keys = ('length', 'thick', 'id')
        sql = "SELECT Round(Sum((tb.Length+{0})*(te.Count))/10^3, 2), tb.Width, te.PriceID " \
              "FROM TBands AS tb INNER JOIN TElems AS te ON tb.UnitPos = te.UnitPos {1} " \
              "GROUP BY tb.Width, te.PriceID".format(add, filtrtpp)
        res = self.db.rs(sql)
        lres = []
        if res:
            for i in res:
                Band = namedtuple('Band', keys)
                lres.append(Band(*i))
        return lres

# -*- coding: utf-8 -*-
__author__ = 'Виноградов А.Г. г.Белгород  август 2015'


class Nomenclature:
    '''Класс работы с номенклатурой'''
    def __init__(self, db):
        try:
            self.db = db
        except:
            return None

    def accbyuid(self, uid=None, tpp=None):
        '''Выводит список аксессуаров
            uid = id единицы измерения
            tpp = TopParentPos ID верхнего хозяина аксессуара
            ID, Название, артикль, ед.изм., кол-во, цена'''
        filtruid = "AND tnn.UnitsID={}".format(uid)
        if uid is None:
            filtruid = ""
        filtrtpp = " AND te.TopParentPos={}".format(tpp)
        if tpp is None:
            filtrtpp = ""
        keys = ('ID', 'Name', 'Article', 'UnitsName', 'Count', 'Price')

        sql = "SELECT ID, [Name], Article, UnitsName, Sum(Cnt/IIf(UnitsID=10,accCnt,1)), Price FROM " \
              "(SELECT tnn.ID, tnn.Name, tnn.Article, tnn.UnitsID, tnn.UnitsName, (te.Count) AS Cnt, tnn.Price, " \
              "(Select (Max(ta1.AccType)-Min(ta1.AccType))+1 from TAccessories AS ta1 where ta1.AccMatID=ta.AccMatID) AS accCnt " \
              "FROM (TAccessories AS ta LEFT JOIN TNNomenclature AS tnn ON ta.AccMatID = tnn.ID) LEFT JOIN TElems AS te " \
              "ON ta.UnitPos = te.UnitPos WHERE tnn.UnitsID<>11 {0}) GROUP BY ID, [Name], Article, UnitsName, Price " \
              "ORDER BY [Name]".format(filtruid+filtrtpp)
        res = self.db.rs(sql)
        dres = []
        for i in res:
            dres.append(dict(zip(keys,i)))
        return dres

    def acclong(self, tpp=None):
        '''Список погонажных комплектующих, таких как сетки
           tpp = TopParentPos ID верхнего хозяина аксессуара
           ID, Название, артикль, ед.изм., длина, кол-во, цена'''
        filtrtpp = ("WHERE te.TopParentPos={}".format(tpp) if tpp else '')
        keys = ('ID', 'Name', 'Article', 'UnitsName', 'Length', 'Count', 'Price')
        sql = "SELECT tnn.ID, tnn.Name, tnn.Article, tnn.UnitsName, te.XUnit/1000, te.Count, tnn.Price FROM TElems AS te " \
              "INNER JOIN TNNomenclature AS tnn ON te.PriceID = tnn.ID WHERE te.FurnType Like '07%' {0} ORDER BY te.Name".format(filtrtpp)
        res = self.db.rs(sql)
        dres = []
        for i in res:
            dres.append(dict(zip(keys,i)))
        return dres

    def matbyuid(self, uid, tpp=None):
        '''Выводит список ID материалов
            uid = id единицы измерения
            tpp = TopParentPos ID верхнего хозяина материала'''
        
        filtrtpp = "{}".format(uid)
        if tpp:
            filtrtpp = "{} AND te.TopParentPos={}".format(uid, tpp)
        
        sql = "SELECT tnn.ID FROM TElems AS te INNER JOIN TNNomenclature AS tnn ON te.PriceID = tnn.ID " \
              "WHERE tnn.UnitsID={} GROUP BY tnn.ID".format(filtrtpp)
        
        res = self.db.rs(sql)
        id = []
        for i in res:
            id.append(i[0])
        return id

    def properties(self, id=0):
        '''Список всех свойств номенклатурной единицы в виде словаря'''
        sql = "SELECT tnp.Ident, Switch([tnp].[TypeID]=1,[tnpv].[DValue],[tnp].[TypeID]=3,[tnpv].[IValue]," \
              "[tnp].[TypeID]=5,[tnpv].[SValue],[tnp].[TypeID]=6,[tnpv].[IValue],[tnp].[TypeID]=7,[tnpv].[DValue]," \
              "[tnp].[TypeID]=11,[tnpv].[IValue],[tnp].[TypeID]=12,[tnpv].[SValue],[tnp].[TypeID]=13,[tnpv].[SValue]," \
              "[tnp].[TypeID]=14,[tnpv].[SValue],[tnp].[TypeID]=15,[tnpv].[SValue],[tnp].[TypeID]=16,[tnpv].[SValue]," \
              "[tnp].[TypeID]=17,[tnpv].[IValue],[tnp].[TypeID]=18,[tnpv].[IValue]) AS val " \
              "FROM TNProperties AS tnp INNER JOIN TNPropertyValues AS tnpv ON tnp.ID = tnpv.PropertyID WHERE tnpv.EntityID={0}" \
              " UNION " \
              "SELECT 'Name', tnn.Name FROM TNNomenclature AS tnn WHERE tnn.ID={0}" \
              " UNION " \
              "SELECT 'Article', tnn.Article FROM TNNomenclature AS tnn WHERE tnn.ID={0}" \
              " UNION " \
              "SELECT 'UnitsID', tnn.UnitsID FROM TNNomenclature AS tnn WHERE tnn.ID={0}" \
              " UNION " \
              "SELECT 'UnitsName', tnn.UnitsName FROM TNNomenclature AS tnn WHERE tnn.ID={0}" \
              " UNION " \
              "SELECT 'Price', tnn.Price FROM TNNomenclature AS tnn WHERE tnn.ID={0}" \
              " UNION " \
              "SELECT 'MatTypeID', tnn.MatTypeID FROM TNNomenclature AS tnn WHERE tnn.ID={0} " \
              "ORDER BY Ident".format(id)
        res = self.db.rs(sql)
        return dict(res)
    
    def property_name(self, id, prop):
        '''Получить значение именованного сво-ва
           id - номер материала
           prop - имя свойства'''
        keys = ('ID', 'Name', 'MatTypeID', 'MatTypeName', 'GroupID', 'GroupName', 'KindID', \
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
        '''Определяем кол-во площадного материала'''
        filtrtpp = " AND te.TopParentPos={}".format(tpp)
        if tpp is None:
            filtrtpp = ""
        sql = "SELECT Sum((XUnit*YUnit*Count)/10^6) AS sqr FROM TElems AS te WHERE te.PriceID={}".format(str(id)+filtrtpp)
        res = self.db.rs(sql)
        return res[0][0]

    def matcount(self, id, tpp=None):
        '''Выводит кол-во материала'''
        cnt = None
        # Проверим, является ли материал аксессуаром
        sql = "SELECT Count(AccMatID) FROM TAccessories WHERE TAccessories.AccMatID={}".format(id)
        res = self.db.rs(sql)
        try:
            if res[0][0]>0:
                cnt = self.accbyuid(id, tpp)
            else:
                prop = self.properties(id)
                unit = prop['UnitsID']
                if unit == '2':               # Квадратные метры
                    cnt = self.sqm(id, tpp)
        except:
            pass

        return cnt

    def bands(self, add=0, tpp=None):
        '''Информация по кромке: длина, толщина торца, ID материала
            add - добавочная длина кромки на торец для отходов
            tpp - ID хозяина кромки'''
        filtrtpp = ("WHERE te.TopParentPos={}".format(tpp) if tpp else '')
        keys = ('Length', 'Thickness', 'ID')
        sql = "SELECT Sum((tb.Length+{0})*(te.Count))/10^3, tb.Width, te.PriceID " \
              "FROM TBands AS tb INNER JOIN TElems AS te ON tb.UnitPos = te.UnitPos {1} " \
              "GROUP BY tb.Width, te.PriceID".format(add, filtrtpp)
        res = self.db.rs(sql)
        dres = []
        for i in res:
            dres.append(dict(zip(keys,i)))
        return dres
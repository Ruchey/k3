# -*- coding: cp1251 -*-
__author__ = 'Виноградов А.Г. г.Белгород'

# Модуль создающий таблицы данных
# t_sheets     Таблица листовыйх материалов
# t_components Таблица комлпектующих
# t_bands      Таблица кромок
# t_profiles   Таблица кусков профилей
# t_sumprof    Таблица профилей
# t_longs      Таблица длиномеров
# 
# 


# import k3
from k3r import *


class Specific:
    """Получение спецификации"""

    def __init__(self, db):
        self.nm = nomenclature.Nomenclature(db)
        self.pf = prof.Profile(db)
        self.bs = base.Base(db)
        self.ln = long.Long(db)

    def t_sheets(self, tpp=None):
        'Таблица листовыйх материалов'
        matid = self.nm.mat_by_uid(2, tpp)
        sh = []
        if matid:
            for i in matid:
                prop = self.nm.properties(i)
                sh.append(prop.id)
                sh.append(prop.name)
                sh.append(prop.price)
                sh.append(prop.unitsname)
                sh.append(prop.pricecoeff)
                sh.append(prop.wastecoeff)
                sh.append(prop.thickness)
                sh.append(prop.density)
                sh.append(round(self.nm.sqm(i), 1))
                sh.append(prop.mattypeid)
            # sh.sort(key=lambda x: [int(xmattypeid),int(xthickness)], reverse=True)
        return sh

    def t_components(self, tpp=None):
        'Таблица комлпектующих'
        comp = []
        uid = [4, 10]  # шт компл.
        for i in uid:
            acclist = self.nm.acc_by_uid(i, tpp)
            if acclist:
                comp.append(acclist)
        accln = self.nm.acc_long(tpp)
        if accln:
            comp.append(accln)
        for i in comp:
            for j in i:
                prop = self.nm.properties(j.get('ID'))
                j['PriceCoeff'] = prop.get('PriceCoeff')
                j['WasteCoeff'] = prop.get('WasteCoeff')
        return comp

    def t_bands(self, add=0, tpp=None):
        'Таблица кромок'
        res = self.nm.bands(add, tpp)
        for i in res:
            prop = self.nm.properties(i.get('ID'))
            i['PriceCoeff'] = prop.get('PriceCoeff')
            i['WasteCoeff'] = prop.get('WasteCoeff')
            i['Name'] = prop.get('Name')
            i['UnitsName'] = prop.get('UnitsName')
            i['Price'] = prop.get('Price')
            i['BandType'] = int(prop.get('BandType'))
            i['Length'] = round(i.get('Length'), 1)
        res.sort(key=lambda x: x['BandType'])
        return res

    def t_profiles(self, tpp=None):
        'Таблица кусков профилей'
        res = self.pf.profiles(tpp)
        dic = []
        keys = ('UnitPos', 'Length', 'ColorID', 'FormType')
        for i in res:
            tmp = dict(zip(keys, i))
            tlms = self.bs.telems(tmp['UnitPos'])
            tmp['PriceID'] = tlms['PriceID']
            prop = self.nm.properties(tmp['PriceID'])
            tmp['Price'] = prop['Price']
            tmp['UnitsName'] = prop['UnitsName']
            tmp['PriceCoeff'] = prop.get('PriceCoeff')
            tmp['WasteCoeff'] = prop.get('WasteCoeff')
            tmp['Article'] = prop.get('Article')
            tmp['Name'] = prop.get('Name')
            tmp['maxLength'] = prop.get('Length')
            tmp['minLength'] = prop.get('minLength')
            tmp['stepcut'] = prop.get('stepcut')
            dic.append(tmp)
        return dic

    def t_sumprof(self, tpp=None):
        'Таблица профилей'
        res = self.pf.total(tpp)
        dic = []
        keys = ('PriceID', 'Length', 'ColorID')
        for i in res:
            tmp = dict(zip(keys, i))
            prop = self.nm.properties(tmp['PriceID'])
            tmp['Price'] = prop['Price']
            tmp['UnitsName'] = prop['UnitsName']
            tmp['PriceCoeff'] = prop.get('PriceCoeff')
            tmp['WasteCoeff'] = prop.get('WasteCoeff')
            tmp['Article'] = prop.get('Article')
            tmp['Name'] = prop.get('Name')
            tmp['maxLength'] = prop.get('Length')
            tmp['minLength'] = prop.get('minLength')
            tmp['stepcut'] = prop.get('stepcut')
            dic.append(tmp)
        return dic

    def t_longs(self, tpp=None):
        'Таблица длиномеров'
        keys = ('UnitPos', 'LongType', 'LongTable', 'LongMatID', 'LongGoodsID')
        dic = []
        res = self.ln.long_list(tpp)
        for i in res:
            tmp = dict(zip(keys, i))
            prop = self.nm.properties(tmp['LongMatID'])
            tngoods = self.bs.tngoods(tmp['LongMatID'])
            tmp['GroupName'] = tngoods['GroupName']

            tmp['Price'] = prop['Price']
            tmp['UnitsName'] = prop['UnitsName']
            tmp['PriceCoeff'] = prop.get('PriceCoeff')
            tmp['WasteCoeff'] = prop.get('WasteCoeff')
            tmp['Article'] = prop.get('Article')
            tmp['Name'] = prop.get('Name')
            tmp['maxLength'] = prop.get('Length')
            tmp['minLength'] = prop.get('minLength')
            tmp['stepcut'] = prop.get('stepcut')
            dic.append(tmp)
        return dic


def start():
    sp = Specific(db)
    for i in sp.t_sheets():
        print(i)
    # for i in sp.t_longs():
    # print(i)


if __name__ == '__main__':

    file = (r'd:\К3\КМ\КМ черновик\46\46.mdb')
    # file = (r'd:\K3\2017\57\57.mdb')
    # file = (r'd:\PKMProjects73\42\42.mdb')

    db = db.DB()
    tmp = db.open(file)  # Подключаемся к базе выгрузки
    if tmp == 'NoFile':
        print('Такой базы данных нет')
        raise SystemExit(1)
    try:
        start()
    except:
        print('Произошла ошибка во время создания отчёта')
    db.close()

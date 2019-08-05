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


#import k3
from MReports import *

class Specific:
    '''Получение спецификации'''
    def __init__(self, db):
        self.nm = Nomenclature(db)
        self.pf = Profile(db)
        self.bs = Base(db)
        self.ln = Longs(db)
    
    def t_sheets(self, tpp=None):
        'Таблица листовыйх материалов'
        matid = self.nm.matbyuid(2, tpp)
        sh = []
        if matid:
            for i in matid:
                dic = {}
                prop = self.nm.properties(i)
                dic['ID'] = i
                dic['Name'] = prop.get('Name')
                dic['Price'] = prop.get('Price')
                dic['UnitsName'] = prop.get('UnitsName')
                dic['PriceCoeff'] = (1 if not prop.get('PriceCoeff') else prop.get('PriceCoeff'))
                dic['WasteCoeff'] = (1 if not prop.get('WasteCoeff') else prop.get('WasteCoeff'))
                dic['Thickness'] = prop.get('Thickness')
                dic['density'] = (780 if not prop.get('density') else prop.get('density'))
                dic['Count'] = round(self.nm.sqm(i), 1)
                dic['MatTypeID'] = prop.get('MatTypeID')
                sh.append(dic)
            sh.sort(key=lambda x: [int(x['MatTypeID']),int(x['Thickness'])], reverse=True)
        return sh
    
    def t_components(self, tpp=None):
        'Таблица комлпектующих'
        comp = []
        uid = [4, 10]    # шт компл.
        for i in uid:
            acclist = self.nm.accbyuid(i, tpp)
            if acclist:
                comp.append(acclist)
        accln = self.nm.acclong(tpp)
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
            i['Length'] = round(i.get('Length'),1)
        res.sort(key=lambda x: x['BandType'])
        return res
    
    def t_profiles(self, tpp=None):
        'Таблица кусков профилей'
        res = self.pf.pflist(tpp)
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
        res = self.pf.sumcount(tpp)
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
        res = self.ln.llist(tpp)
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
    #for i in sp.t_longs():
        #print(i)

if __name__ == '__main__':
    
    
    file = (r'd:\К3\КМ\КМ черновик\46\46.mdb')
    #file = (r'd:\K3\2017\57\57.mdb')
    #file = (r'd:\PKMProjects73\42\42.mdb')
    
    db = DB()
    tmp = db.connect(file)    # Подключаемся к базе выгрузки
    if tmp == 'NoFile':
        print('Такой базы данных нет')
        raise SystemExit(1)
    try:
        start()
    except:
        print('Произошла ошибка во время создания отчёта')
    db.disconnect()
    

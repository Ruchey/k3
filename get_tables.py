# -*- coding: cp1251 -*-
__author__ = '���������� �.�. �.��������'

# ������ ��������� ������� ������
# t_sheets     ������� ��������� ����������
# t_components ������� �������������
# t_bands      ������� ������
# t_profiles   ������� ������ ��������
# t_sumprof    ������� ��������
# t_longs      ������� ����������
# 
# 


# import k3
from k3r import *


class Specific:
    """��������� ������������"""

    def __init__(self, db):
        self.nm = nomenclature.Nomenclature(db)
        self.pf = prof.Profile(db)
        self.bs = base.Base(db)
        self.ln = long.Long(db)

    def t_sheets(self, tpp=None):
        '������� ��������� ����������'
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
        '������� �������������'
        comp = []
        uid = [4, 10]  # �� �����.
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
        '������� ������'
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
        '������� ������ ��������'
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
        '������� ��������'
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
        '������� ����������'
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

    file = (r'd:\�3\��\�� ��������\46\46.mdb')
    # file = (r'd:\K3\2017\57\57.mdb')
    # file = (r'd:\PKMProjects73\42\42.mdb')

    db = db.DB()
    tmp = db.open(file)  # ������������ � ���� ��������
    if tmp == 'NoFile':
        print('����� ���� ������ ���')
        raise SystemExit(1)
    try:
        start()
    except:
        print('��������� ������ �� ����� �������� ������')
    db.close()

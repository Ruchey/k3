# -*- coding: cp1251 -*-
__author__ = 'Виноградов А.Г. г.Белгород 2018'
# Формирование бланка заказа на распил для фабрики SkyLine

import k3
from MReports import *
import itertools
import operator
import os

LISTKEY = ('num', 'plength', 'pwidth', 'length', 'width', 'band_x04', 'band_y04', 'band_x2', 'band_y2', 'dir')


def start(tpp=None):
    # 'Формируем данные'
    matid = nm.matbyuid(2, tpp)
    listDSP = []
    listDVP = []
    if matid:
        for i in matid:
            dic = {}
            prop = nm.properties(i)
            dic['ID'] = i
            dic['Name'] = prop.get('Name', '')
            dic['Thickness'] = int(prop.get('Thickness', 1))
            dic['MatTypeID'] = int(prop.get('MatTypeID', 1))
            dic['GabX'] = int(prop.get('GabX', 1))
            dic['GabY'] = int(prop.get('GabY', 1))
            if dic['MatTypeID'] == 128:
                listDSP.append(dic)
                listDSP.sort(key=lambda x: int(x['Thickness']), reverse=True)
            elif dic['MatTypeID'] == 37:
                listDVP.append(dic)
                listDVP.sort(key=lambda x: int(x['Thickness']), reverse=True)

        if listDSP:
            rep_pan('ДСП', listDSP, tpp)

        if listDVP:
            rep_pan('ДВП', listDVP, tpp)


def group_list_panels(lst_dpars, listkey=LISTKEY):
    """Группирует словари в списке lst_dpars
    по совпадающим ключевым значениям из списка listkey
    Возвращает список словарей из двух элементов '_key' и 'data'
    '_key' содержит кортеж одинаковых значений в порядке ключей представленных в listkey
    для словарей представленных в data
    >>> [{'_key': ('полка', 700, 400, 0, 1, 1, 1, 1),
...   'data': [{'UnitPos': 1, 'holes': [1, 2, 3], 'другие парметры': 5},
...            {'UnitPos': 4, 'holes': [1, 2, 3], 'другие парметры': 5},
...            {'UnitPos': 6, 'holes': [1, 2, 3], 'другие парметры': 5},
...            {'UnitPos': 8, 'holes': [1, 2, 3], 'другие парметры': 5},
...            {'UnitPos': 10, 'holes': [1, 2, 3], 'другие парметры': 5}]},
...  {'_key': ('стенка', 700, 400, 0, 1, 1, 1, 1),
...   'data': [{'UnitPos': 3, 'holes': [1, 2, 3], 'другие парметры': 5}]},
...  ]
    """

    def _groupid_drop(d):
        for nm in listkey:
            del d[nm]
        return d

    lst_dpars.sort(key=lambda x: [x[k] for k in listkey])

    return [{'_key': i, 'data': list(map(_groupid_drop, grp))} for i, grp in
            itertools.groupby(lst_dpars, operator.itemgetter(*map(lambda nm: nm, listkey)))
            ]


def list_pan(matid, tpp=None):
    '''Получаем список деталей панелей'''
    idPan = pn.list_panels(matid, tpp)
    listPan = []
    listPan2 = []
    for i in idPan:
        dic = {}
        cb = ''
        if pn.form(i) == 0:
            telems = bs.telems(i)
            dic['UnitPos'] = i
            dic['num'] = telems['CommonPos']
            dic['name'] = telems['Name']
            dic['plength'] = pn.planelength(i)
            dic['pwidth'] = pn.planewidth(i)

            dic['length'] = pn.length(i)
            dic['width'] = pn.width(i)
            pdir = pn.dir(i)
            dic['dir'] = 'N' if pdir == -1 else 'Y'
            dic['cnt'] = telems['Count']

            bb = pn.band_b(i)
            bc = pn.band_c(i)
            bd = pn.band_d(i)
            be = pn.band_e(i)

            dic['band_x04'] = int(0 < be['ThickBand'] < 2) + int(0 < bd['ThickBand'] < 2)
            dic['band_y04'] = int(0 < bb['ThickBand'] < 2) + int(0 < bc['ThickBand'] < 2)
            dic['band_x2'] = int(0 < be['ThickBand'] == 2) + int(0 < bd['ThickBand'] == 2)
            dic['band_y2'] = int(0 < bb['ThickBand'] == 2) + int(0 < bc['ThickBand'] == 2)

            if (45 < pdir <= 135) or (225 < pdir <= 315):
                dic['band_x04'], dic['band_y04'] = dic['band_y04'], dic['band_x04']
                dic['band_x2'], dic['band_y2'] = dic['band_y2'], dic['band_x2']

            dic['plength'] += fuga * (dic['band_y04'] + dic['band_y2'])  # добавляем к размеру прифуговку
            dic['pwidth'] += fuga * (dic['band_x04'] + dic['band_x2'])  # добавляем к размеру прифуговку

            listPan.append(dic)

    for i in group_list_panels(listPan):
        dic = dict(zip(LISTKEY, i['_key']))
        cnt = 0
        for j in i['data']:
            cnt += j['cnt']
        dic.update({'cnt': cnt})
        listPan2.append(dic)

    listPan2.sort(key=operator.itemgetter('length', 'width'), reverse=True)
    return listPan2


def newsheet(xl, name):
    'Создаём новый лист с именем'
    xl.sheetorient = xl.xlconst['xlPortrait']
    xl.rightmargin = 0.6
    xl.leftmargin = 0.6
    xl.bottommargin = 1.6
    xl.topmargin = 1.6
    xl.centerhorizontally = 0
    xl.displayzeros = False
    xl.new_sheet(name)
    return


def rep_pan(name, mat, tpp=None):
    'Вывод деталей панелей'
    fnum = '0,00'  # числовой формат
    fgen = ''  # общий формат
    txt = '@'  # текстовый формат
    # 'Выводим таблицу листовых материалов'
    for imat in mat:
        xl = ExcelDoc()
        xl.Excel.Visible = 0
        name = imat['Name'].replace('/', '_')
        newsheet(xl, name)
        formarsheet = [imat['GabX'], imat['GabY']]
        row = draw_table(xl, 1, 'Виноградов А.Г. 8920-206-83-98', imat['Name'], 'В цвет', formarsheet)
        rw1 = row
        index = 1
        for l in list_pan(imat['ID'], tpp):
            val = (l['num'], l['length'], '', 'x', l['width'], '', l['cnt'], l['dir'])
            val2 = (l['band_x04'], l['band_x2'], '', l['band_y04'], l['band_y2'], '', '', \
                    '=(R[-1]C[-7]*RC[-7]*R[-1]C[-2]+R[-1]C[-4]*RC[-4]*R[-1]C[-2])', \
                    '=(R[-1]C[-8]*R[-1]C[-5]*R[-1]C[-3])/10^6', \
                    '=(R[-1]C[-9]*RC[-8]*R[-1]C[-4]+R[-1]C[-6]*RC[-5]*R[-1]C[-4])')
            xl.putval(row, 1, val)
            xl.mergecells(row, 2, 2, hg=1, halign='c', valign='c', wrap='f')
            xl.mergecells(row, 5, 2, hg=1, halign='c', valign='c', wrap='f')
            xl.mergecells(row, 7, 1, hg=2, halign='c', valign='c', wrap='f')
            xl.mergecells(row, 8, 1, hg=2, halign='c', valign='c', wrap='f')
            xl.rowsize(row, 18)
            row += 1
            xl.putval(row, 2, val2)
            xl.mergecells(row - 1, 1, 1, hg=2, halign='r', valign='c', wrap='f')
            xl.mergecells(row - 1, 9, 1, hg=2, halign='c', valign='c', wrap='f')
            xl.mergecells(row - 1, 10, 1, hg=2, halign='c', valign='c', wrap='f', nf=fnum)
            xl.mergecells(row - 1, 11, 1, hg=2, halign='c', valign='c', wrap='f')
            xl.txtformat(row - 1, 2, halign='ccccc', valign='c', bold='t')
            if index % 2 != 0:
                xl.paintcells(row - 1, 1, ln=11, tc=1, ink=2, tas=1)
                xl.paintcells(row, 2, ln=5, tc=1, ink=2, tas=2)
            xl.txtformat(row, 2, halign='ccccc', valign='c')
            row += 1
            index += 1
        gh = row - rw1
        xl.gridset(tc=2)
        xl.grid(rw1, 1, 11, gh, 'lrudvh')

        # Вывоим данные для раскроя
        newsheet(xl, 'CutRite')
        sp_tab = [13.3, 12, 9, 9, 9, 10]
        xl.columnsize(1, sp_tab)
        xl.putval(1, 1, 'Учтена прифуговка {}мм'.format(fuga))
        val = ['Обозначение', 'Материал', 'Длина', 'Ширина', 'Кол-во', 'Структура']
        xl.putval(3, 1, val)
        row = 4
        for l in list_pan(imat['ID'], tpp):
            val = (l['num'], '', l['plength'], l['pwidth'], l['cnt'], l['dir'])
            xl.putval(row, 1, val)
            row += 1
        xl.movesheet(2, 1)
        xl.save(os.path.join(projreppath, 'SkyLine_{}'.format(name)))
        xl.Excel.Visible = 1


def draw_table(xl, rw1, customer, mat, bands, fs):
    'Форматируем таблицу'
    fnum = '0,000'  # числовой формат
    xl.putval(rw1, 1, ['Заказчик', '', '', '', '', customer])
    xl.putval(rw1 + 1, 1, ['Цвет и код ЛДСП', '', '', '', '', mat])
    xl.putval(rw1 + 2, 1, ['Цвет и код кромки', '', '', '', '', bands])
    xl.mergecells(rw1, 1, 5)
    xl.mergecells(rw1, 6, 7)
    xl.mergecells(rw1 + 1, 1, 5)
    xl.mergecells(rw1 + 1, 6, 7)
    xl.mergecells(rw1 + 2, 1, 5)
    xl.mergecells(rw1 + 2, 6, 7)
    xl.txtformat(rw1, 1, halign='lllllll', valign='c', bold='tf')
    xl.txtformat(rw1 + 1, 1, halign='lllllll', valign='c', bold='tf')
    xl.txtformat(rw1 + 2, 1, halign='lllllll', valign='c', bold='tf')
    xl.gridset(tc=2)
    xl.grid(rw1, 1, 12, 3, 'lrudvh')
    rw1 += 4
    s_sheet = (fs[0] * fs[1]) / 10 ** 6
    head = ['№', 'Длина', '', 'x', 'Ширина', '', 'кол-во', 'текстура', 'Квадратура деталей', \
            '=SUM(R[4]C:R[1000]C)', 'Кол-во листов формата {0}'.format(str(fs[0]) + 'x' + str(fs[1])),
            '=RC[-2]*1.2/{0}'.format(s_sheet)]
    head2 = ['', 'Кол-во сторон кромки', '', '', '', '', '', '', 'Кол-во кромки 0,4мм', '=SUM(C[-1])/10^3',
             'Кол-во кромки 2мм', '=SUM(C[-1])/10^3']
    head3 = ['', '0,4мм', '2мм', 'x', '0,4мм', '2мм', '', '', 'Купить кромки 0,4мм', '=R[-1]C*1.1', 'Купить кромки 2мм',
             '=R[-1]C*1.1']
    sp_tab = [2.4, 5.71, 5.71, 1.3, 5.71, 5.71, 3, 3, 21, 7.3, 21, 7.3]
    xl.rowsize(rw1, 30)
    xl.mergecells(rw1, 2, 2)
    xl.mergecells(rw1, 5, 2)
    xl.mergecells(rw1 + 1, 2, 5)
    xl.mergecells(rw1, 7, 1, 3)
    xl.mergecells(rw1, 8, 1, 3)
    xl.columnsize(1, sp_tab)
    xl.putval(rw1, 1, head)
    xl.putval(rw1 + 1, 1, head2)
    xl.putval(rw1 + 2, 1, head3)

    xl.txtformat(rw1, 1, halign='c', valign='c', wrap='t', ort=[0, 0, 0, 0, 0, 0, 90, 90, 0, 0, 0, 0], fsz=[11],
                 nf=[fnum])
    xl.txtformat(rw1 + 1, 10, halign='c', nf=[fnum])
    xl.txtformat(rw1 + 2, 1, halign='cccccccclclc', valign='c', nf=[fnum], bold='fftfftfft')
    xl.gridset(tc=2)
    xl.grid(rw1, 1, 12, 3, 'lrudvh')
    rw1 += 4
    return rw1


if __name__ == '__main__':

    # file = r'd:\PKMProjects73\1\1.mdb'
    # projreppath = r'd:\PKMProjects73\1'
    # fuga = 1

    params = k3.getpar()
    file = params[0]
    projreppath = params[1]
    fuga = params[3]

    db = DB()
    tmp = db.connect(file)  # Подключаемся к базе выгрузки
    if tmp == 'NoFile':
        raise Exception('Ошибка подключения к базе данных')
    try:
        nm = Nomenclature(db)
        bs = Base(db)
        pn = Panel(db)
        start()
        params[2].value = 1
    except:
        raise Exception('Произошла ошибка во время создания отчёта')
    db.disconnect()

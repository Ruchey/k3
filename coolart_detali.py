# -*- coding: utf-8 -*-

import k3Report
import os


def newsheet(name):
    'Создаём новый лист с именем'

    xl.sheetorient = xl.PORTRAIT
    xl.rightmargin = 0.6
    xl.leftmargin = 0.6
    xl.bottommargin = 1.6
    xl.topmargin = 1.6
    xl.centerhorizontally = 0
    xl.displayzeros = False
    xl.newsheet(name)
    return



def list_pan(matid, tpp=None):
    '''Получаем список деталей панелей'''

    idPans = pn.list_panels(matid, tpp)
    listPan = []
    for idpan in idPans:
        pans = {}
        if pn.form(idpan) == 0:
            telems = bs.telems(idpan)
            pans['name'] = telems['Name']
            pans['cnt'] = telems['Count']
            pans['length'] = pn.planelength(idpan)
            pans['width'] = pn.planewidth(idpan)
            pdir = pn.dir(idpan)


def rep_pan(name, mat, tpp=None):

    newsheet(name)
    row = 1
    for i, imat in enumerate(mat):
        for l in list_pan(imat['ID'], tpp):
            val = (i+1, l['name'], imat['Thickness'], imat['Article'], l['length'], l['width'], l['cnt'], \
                   l['band_x1'], l['band_x2'], l['band_y1'], l['band_y2'], l['comment']
                   )
            xl.putval(row, 1, val)
            row += 1



def makereport(tpp=None):
    # 'Формируем данные'
    matid = nm.matbyuid(2, tpp)
    if not matid:
        return
    listmat = []
    listDSP = []
    listMDF = []
    listDVP = []
    listGLS = []
    for i in matid:
        prop = nm.properties(i)
        if prop.mattypeid == 128:
            listDSP.append(prop)
        elif prop.mattypeid == 64:
            listMDF.append(prop)
        elif prop.mattypeid == 37:
            listDVP.append(prop)
        elif prop.mattypeid in [48, 99]:
            listGLS.append(prop)
        else:
            listmat.append(prop)

    if listDSP:
        listDSP.sort(key=lambda x: x.thickness, reverse=True)
        rep_pan('ДСП', listDSP, tpp)

    if listMDF:
        listMDF.sort(key=lambda x: x.thickness, reverse=True)
        rep_pan('МДФ', listMDF, tpp)

    if listDVP:
        listDVP.sort(key=lambda x: x.thickness, reverse=True)
        rep_pan('ДВП', listDVP, tpp)

    if listGLS:
        listGLS.sort(key=lambda x: x.thickness, reverse=True)
        rep_pan('Стекло', listGLS, tpp)

    if listmat:
        listmat.sort(key=lambda x: [x.mattypeid,x.thickness], reverse=True)
        rep_pan('Прочее', listmat, tpp)


def start(fileDB, projreppath):
    
    global nm, bs, pn, xl

    # Подключаемся к базе выгрузки
    db = k3Report.db.DB(fileDB)
    if db == 'NoFile':
        raise Exception('Ошибка подключения к базе данных')

    nm = k3Report.nomenclature.Nomenclature(db)
    bs = k3Report.base.Base(db)
    pn = k3Report.panel.Panel(db)

    # Создаём объект документа
    xl = k3Report.doc.DocOpenpyxl()
    makereport()
    db.close()
    xl.save(os.path.join(projreppath, 'Деталировка.xlsx'))



if __name__=='__main__':

    fileDB = r'd:\К3\Самара\Самара черновик\1\1.mdb'
    projreppath = r'd:\К3\Самара\Самара черновик\1\Reports'
    start(fileDB, projreppath)

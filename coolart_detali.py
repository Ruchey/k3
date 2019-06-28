# -*- coding: utf-8 -*-

import k3Report
import os

from collections import namedtuple

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
    keys = ('name', 'thickness', 'article', 'length', 'width', 'cnt',
            'band_x1', 'band_x2', 'band_y1', 'band_y2', 'num', 'note')
    for idpan in idPans:
        pans = namedtuple('pans', keys)
        if pn.form(idpan) == 0:
            telems = bs.telems(idpan)
            matprop = nm.properties(matid)
            pans.name = telems.name
            pans.thickness = matprop.thickness
            pans.article = matprop.article
            pans.length = pn.planelength(idpan)
            pans.width = pn.planewidth(idpan)
            pans.cnt = telems.count
            pans.band_x1 = pn.band_x1(idpan).thickband
            pans.band_x2 = pn.band_x2(idpan).thickband
            pans.band_y1 = pn.band_y1(idpan).thickband
            pans.band_y2 = pn.band_y2(idpan).thickband
            pans.num = telems.commonpos
            note_curvepath = 'Фигурная' if pn.curvepath(idpan) else ''
            slotx = list(map('{0.beg}ш{0.width}г{0.depth}'.format,
                             pn.slots_x_par(idpan)))
            sloty = list(map('{0.beg}ш{0.width}г{0.depth}'.format,
                             pn.slots_y_par(idpan)))
            note_slotx = "Паз{0} по X {1}".format("ы" if len(slotx) > 1 else "",
                                                  "; ".join(slotx)) if slotx else ""
            note_sloty = "Паз{0} по Y {1}".format("ы" if len(sloty) > 1 else "",
                                                  "; ".join(sloty)) if sloty else ""
            notes = [note_curvepath, note_slotx, note_sloty]
            pans.note = ". ".join(list(filter(None, notes)))
            pdir = pn.dir(idpan)
            listPan.append(pans)
    return listPan


def rep_pan(name, mat, tpp=None):

    newsheet(name)
    row = 1
    for i, imat in enumerate(mat):
        pans = list_pan(imat.id, tpp)
        for l in pans:
            val = (l.num, l.name, imat.thickness, imat.name, l.length, l.width, l.cnt, \
                   l.band_x1, l.band_x2, l.band_y1, l.band_y2, l.note
                   )
            row = xl.putval(row, 1, val)
            



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
    fuga = 0.5
    start(fileDB, projreppath)

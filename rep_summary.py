# -*- coding: utf-8 -*-

import k3Report
import os
import openpyxl
from collections import namedtuple


class Report:
    
    def __init__(self, db):
        
        self.nm = k3Report.nomenclature.Nomenclature(db)
        self.bs = k3Report.base.Base(db)
        self.pn = k3Report.panel.Panel(db)
        self.xl = k3Report.doc.DocOpenpyxl()

    def newsheet(self, name):
        """Создаём новый лист с именем"""
    
        self.xl.sheetorient = self.xl.PORTRAIT
        self.xl.rightmargin = 0.6
        self.xl.leftmargin = 0.6
        self.xl.bottommargin = 1.6
        self.xl.topmargin = 1.6
        self.xl.centerhorizontally = 0
        self.xl.displayzeros = False
        self.xl.newsheet(name)
        return
    
    def totalmat(self, matid, tpp=None):
        """Получаем общее кол-во материала"""
        
        Mat = namedtuple('Mat', 'name unit cnt')
        mat_name = self.nm.property_name(matid, 'Name')
        mat_unit = self.nm.property_name(matid, 'UnitsName')
        mat_cnt = self.nm.matcount(matid, tpp)
        mat = Mat(mat_name, mat_unit, mat_cnt)
        list_bands = self.pn.total_bands_to_panels(self.pn.list_panels(matid, tpp))
        bands = []
        for band in list_bands:
            Band = namedtuple('Band', 'name unit length')
            prop = self.nm.properties(band.priceid)
            band_name = prop.name
            band_unit = prop.unitsname
            band_length = band.length
            bands.append(Band(band_name, band_unit, band_length))
        return (mat, bands)

    def list_pan(self, matid, tpp=None):
        """Получаем список деталей панелей"""
    
        idPans = self.pn.list_panels(matid, tpp)
        listPan = []
        keys = ('cpos', 'name', 'length', 'width', 'cnt',
                'band_x1', 'band_x2', 'band_y1', 'band_y2', 'note')
        for idpan in idPans:
            Pans = namedtuple('Pans', keys)
            if self.pn.form(idpan) == 0:
                telems = self.bs.telems(idpan)
                matprop = self.nm.properties(matid)
                name = telems.name
                thickness = matprop.thickness
                length = self.pn.length(idpan)
                width = self.pn.width(idpan)
                cnt = telems.count
                band_x1 = self.pn.band_x1(idpan).id
                band_x2 = self.pn.band_x2(idpan).id
                band_y1 = self.pn.band_y1(idpan).id
                band_y2 = self.pn.band_y2(idpan).id
                cpos = telems.commonpos
                pdir = self.pn.dir(idpan)
                note_curvepath = 'Фигурная' if self.pn.curvepath(idpan) else ''
                slotx = list(map('{0.beg}ш{0.width}г{0.depth}'.format,
                                 self.pn.slots_x_par(idpan)))
                sloty = list(map('{0.beg}ш{0.width}г{0.depth}'.format,
                                 self.pn.slots_y_par(idpan)))
                if (45<pdir<=135) or (225<pdir<=315):
                    slotx, sloty = sloty, slotx
                note_slotx = "Паз{0} X {1}".format("ы" if len(slotx) > 1 else "",
                                                      "; ".join(slotx)) if slotx else ""
                note_sloty = "Паз{0} Y {1}".format("ы" if len(sloty) > 1 else "",
                                                      "; ".join(sloty)) if sloty else ""
                notes = [note_curvepath, note_slotx, note_sloty]
                note = ". ".join(list(filter(None, notes)))
                pans = Pans(cpos, name, length, width, cnt, band_x1,
                            band_x2, band_y1, band_y2, note
                            )
                listPan.append(pans)
        newlist = k3Report.utils.groupbykey(listPan, 'cpos', 'cnt')
        return newlist


def start(fileDB, projreppath, project):

    # Подключаемся к базе выгрузки
    db = k3Report.db.DB()
    db.open(fileDB)

    rep = Report(db)
    db.close()
    rep.makereport()
    file = os.path.join(projreppath, '{}.xlsx'.format(project))
    rep.xl.save(file)
    os.startfile(file)
    return True


if __name__=='__main__':

    fileDB = r'd:\К3\Самара\Самара черновик\1\1.mdb'
    projreppath = r'd:\К3\Самара\Самара черновик\1\Reports'
    project = "Общий отчёт"
    start(fileDB, projreppath, project)
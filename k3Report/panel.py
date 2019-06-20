# -*- coding: utf-8 -*-

import itertools
import math

from collections import namedtuple


__author__ = 'Виноградов А.Г. г.Белгород  август 2015'


def float_int(num):
    'Убираем нули после точки, если число целое'

    if type(num) != 'float':
        return num

    if num.is_integer():
        return int(num)
    else:
        return round(num, 1)


class Panel:
    def __init__(self, db):
        '''Информация о панели'''
        try:
            self.db = db
        except:
            return None


    def list_panels(self, matid=None, tpp=None):
        '''Получить список панелей по материалу
           matid - id материала из которого сделана панель
           tpp - главный родитель панели
        '''
        
        filtrid = "AND tnn.ID={}".format(matid) if matid else ""
        filtrtpp = "AND te.TopParentPos={}".format(tpp) if tpp else ""
        where = " ".join(["te.FurnType Like '01%'", filtrid, filtrtpp])
        
        sql = "SELECT te.UnitPos FROM TElems AS te INNER JOIN TNNomenclature AS tnn "\
              "ON te.PriceID = tnn.ID WHERE {0}".format(where)
        res = self.db.rs(sql)
        id = []
        if len(res) > 0:
            for i in res:
                id.append(i[0])
        return id


    def xunit(self, unitpos):
        '''Размер панели вдоль X'''
        
        sql = "SELECT xunit FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return float_int(res[0][0])
        else:
            return 0


    def yunit(self, unitpos):
        '''Размер панели вдоль Y'''
        
        sql = "SELECT yunit FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return float_int(res[0][0])
        else:
            return 0


    def zunit(self, unitpos):
        '''Размер панели вдоль Z'''
        
        sql = "SELECT zunit FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return float_int(res[0][0])
        else:
            return 0


    def cnt(self, unitpos):
        '''Количество панелей'''
        
        sql = "SELECT count FROM TElems WHERE unitpos = {}".format(unitpos)
        return self.db.rs(sql)


    def length(self, unitpos):
        '''Длина с кромкой'''
        
        sql = "SELECT length FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return float_int(res[0][0])
        else:
            return 0


    def width(self, unitpos):
        '''Ширина с кромкой'''

        sql = "SELECT width FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return float_int(res[0][0])
        else:
            return 0


    def planelength(self, unitpos):
        '''Длина без кромки'''

        sql = "SELECT planelength FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return float_int(res[0][0])
        else:
            return 0


    def planewidth(self, unitpos):
        '''Ширина без кромки'''
        
        sql = "SELECT planewidth FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return float_int(res[0][0])
        else:
            return 0


    def thickness(self, unitpos):
        '''Толщина панели'''
        
        sql = "SELECT thickness FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return float_int(res[0][0])
        else:
            return 0


    def dir(self, unitpos):
        '''Направление текстуры в градусах'''
        
        sql = "SELECT dir FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return float_int(res[0][0])
        else:
            return 0


    def curvepath(self, unitpos):
        '''Панель отличается от прямоугольности (скос, например)'''
        
        sql = "SELECT curvepath FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return False


    def form(self, unitpos):
        '''Форма панели
        0 - линейная
        1 - дуга по хорде
        2 - два отрезка и дуга
        '''
        
        sql = "SELECT formtype FROM TPanels WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return 0


    def slots(self, unitpos):
        '''Наличие пропилов
        True - есть пропилы
        False - нет пропилов
        '''
        sql = "SELECT Count(panelpos) FROM TSlots WHERE panelpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return bool(res[0][0])
        else:
            return False
    
    
    def slots_x(self, unitpos):
        '''Количество пропилов вдоль длины'''
        
        sql = "SELECT Count(PanelPos) FROM TSlots WHERE ROUND(BegY)=ROUND(EndY) AND PanelPos = {0}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return 0


    def slots_y(self, unitpos):
        '''Количество пропилов поперёк длины'''
        
        sql = "SELECT Count(PanelPos) FROM TSlots WHERE ROUND(BegX)=ROUND(EndX) AND PanelPos = {0}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return 0


    def slots_x_len(self, unitpos):
        '''Суммарная длина пропилов вдоль длины'''
        
        sql = "SELECT ROUND(SUM(ABS(EndX-BegX))/1000,2) FROM TSlots WHERE ROUND(BegY)=ROUND(EndY) AND PanelPos = {0}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return 0


    def slots_y_len(self, unitpos):
        '''Суммарная длина пропилов поперёк длины'''
        
        sql = "SELECT ROUND(SUM(ABS(EndY-BegY))/1000,2) FROM TSlots WHERE ROUND(BegX)=ROUND(EndX) AND PanelPos = {0}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return 0


    def butts(self, unitpos):
        '''Наличие обработки торцов
        True - есть обработка
        False - нет обработки
        '''
        sql = "SELECT Count(unitpos) FROM TButts WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return bool(res[0][0])
        else:
            return False


    def frezerovka(self, unitpos):
        '''Список фрезеровок данной панели'''
        
        sql = "SELECT attrstring FROM TAttributes WHERE name='frezerovka' AND unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return tuple(itertools.chain.from_iterable(res))
        else:
            return ()


    def priceid(self, unitpos):
        '''ID элемента в номенклатурном справочнике'''
        
        sql = "SELECT priceid FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return 0


    def name(self, unitpos):
        '''Название панели'''
        
        sql = "SELECT name FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return ''


    def data(self, unitpos):
        '''Примечание панели'''
        
        sql = "SELECT data FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return ''


    def cmpos(self, unitpos):
        '''Пользовательский номер элемента'''
        
        sql = "SELECT commonpos FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return 0


    def ppos(self, unitpos):
        '''Номер родителя элемента'''
        
        sql = "SELECT parentpos FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return 0


    def tpos(self, unitpos):
        '''Номер верхнего родителя элемента'''
        
        sql = "SELECT topparentpos FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return 0


    def furntype(self, unitpos):
        '''Тип мебельного элемента'''
        
        sql = "SELECT furntype FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return res[0][0]
        else:
            return '000000'


    def sumcost(self, unitpos):
        '''Стоимость элемента'''

        sql = "SELECT sumcost FROM TElems WHERE unitpos = {}".format(unitpos)
        res = self.db.rs(sql)
        if len(res) > 0:
            return float_int(res[0][0])
        else:
            return 0


    def svw(self, unitpos):
        '''Возвращает именованный кортеж
        s - Площадь кв.м
        v - объём куб.м
        w - вес кг
        '''
        svw = namedtuple('svw', 's v w')
        id = self.priceid(unitpos)
        sql = "SELECT tnpv.DValue FROM TNProperties AS tnp LEFT JOIN TNPropertyValues AS tnpv "\
              "ON tnp.ID = tnpv.PropertyID WHERE (((tnp.Ident)='density') AND ((tnpv.EntityID)={}))".format(id)
        res = self.db.rs(sql)
        if len(res) > 0:
            density = int(res[0][0])
            s = self.xunit(unitpos)*self.yunit(unitpos)/1000000
            v = s*self.zunit(unitpos)/1000
            w = v*density
            return svw(round(s,2), round(v,4), round(w,2))
        else:
            return svw((0, 0, 0))


    def radius(self, unitpos):
        '''Возвращает именованный кортеж
        rad - радиус панели
        axis - ось гиба OX или OY
        '''
        ax = {1: 'OX', 2: 'OY'}
        rad = namedtuple('rad', 'rad axis')
        radius = 0
        bend = 0            # ось гиба: 1 - OX 2 - OY
        form = self.form(unitpos)  # узнаём форму панели

        if form == 1:           # дуга по хорде
            sql = "SELECT abs(tpm1.NumValue) AS Caving, tpm2.NumValue AS Chord, tpm3.NumValue AS Bend "\
                  "FROM TParams AS tpm1, TParams AS tpm2, TParams AS tpm3 "\
                  "WHERE (((tpm1.UnitPos)={0}) AND ((tpm1.ParamName)='ArcChord.Caving') AND ((tpm2.UnitPos)={0}) AND "\
                  "((tpm2.ParamName)='ArcChord.Chord') AND ((tpm3.UnitPos)={0}) AND ((tpm3.ParamName)='BendAxis'))".format(unitpos)
            
            res = self.db.rs(sql)
            caving = res[0][0]
            chord = res[0][1]
            bend = ax[res[0][2]]
            alfa = 2*math.atan((2*caving)/chord)
            L = chord*(alfa/math.sin(alfa))
            D = L/alfa
            radius = D/2

        if form == 2:           # Два отрезка и дуга
            sql = "SELECT tpm1.NumValue, tpm2.NumValue AS Bend FROM TParams AS tpm1, TParams AS tpm2 "\
                  "WHERE (tpm1.UnitPos={0} AND tpm1.ParamName='LinesArc.R') AND (tpm2.UnitPos={0} "\
                  "AND tpm2.ParamName='BendAxis')".format(unitpos)
            
            res = self.db.rs(sql)
            radius = res[0][0]
            bend = ax[res[0][1]]

        return rad(round(radius,1), bend)


    def band_side(self, unitpos, IDLine, IDPoly=1):
        '''информация о кромке торца
        id - id кромки из номенклатуры
        name - название кромки
        width - ширина кромки
        thickpan - толщина панели
        thickband - ширина кромки
        count - кол-во кромки
        '''
        band = namedtuple('band', ['id', 'name', 'width', 'thickpan', 'thickband', 'count'])
        b_priceID = 0       # id кромки в номенклатуре
        b_name = ''         # название кромки
        b_width = 0         # ширина кромки
        b_thickPan = 0      # толщина панели
        b_thick = 0         # толщина кромки
        b_cnt = 0

        # определим тип панели
        # PanPolyType: 1 - полилиния 2 - прямоугольная 3 - четырёхугольная
        sql = "SELECT tpm.NumValue FROM TParams AS tpm WHERE (tpm.unitpos = {} AND tpm.ParamName='PanPolyType' AND tpm.HoldTable='TPanels')".format(unitpos)
        res = self.db.rs(sql)
        if not res:
            return band(0, '', 0, 0, 0, 0)
        PolyType = int(res[0][0])
        if (PolyType == 2) or (PolyType == 3):
            sql = "SELECT bu.BandUnitPos, tn.ID, tn.Name, Count FROM (SELECT idp.NumValue AS IDPoly, idl.NumValue AS IDLine, idb.NumValue AS BandUnitPos "\
                  "FROM (TParams idp LEFT OUTER JOIN TParams idl ON idp.UnitPos=idl.UnitPos AND idp.HoldTable=idl.HoldTable AND idp.Hold1=idl.Hold1 "\
                  "AND idp.Hold3=idl.Hold3) LEFT OUTER JOIN TParams idb ON idp.UnitPos=idb.UnitPos AND idp.HoldTable=idb.HoldTable AND idp.Hold1=idb.Hold1 "\
                  "AND idp.Hold3=idb.Hold3 WHERE idp.UnitPos={0} AND idp.HoldTable='TPaths' AND idp.Hold1=1 AND idp.ParamName='IDPoly' AND idl.ParamName='IDLine' "\
                  "AND idb.ParamName='BandUnitPos') bu, TNNomenclature tn, TElems te WHERE bu.BandUnitPos=te.UnitPos AND te.PriceID=tn.ID AND IDPoly={1} AND IDLine={2}".format(unitpos, IDPoly, IDLine)
            res = self.db.rs(sql)
            if res:
                bandID = (res[0][0])
                b_priceID = (res[0][1])
                b_name = (res[0][2])
                b_cnt = (res[0][3])
                sql = "SELECT Width FROM TBands WHERE UnitPos={}".format(bandID)
                res = self.db.rs(sql)
                if res:
                    b_thickPan = (res[0][0])
                sql = "SELECT tnpv.DValue FROM TNPropertyValues AS tnpv WHERE tnpv.PropertyID=21 AND tnpv.EntityID={}".format(b_priceID)
                res = self.db.rs(sql)
                if res:
                    b_width = (res[0][0])
                sql = "SELECT tnpv.DValue FROM TNPropertyValues AS tnpv WHERE tnpv.PropertyID=10 AND tnpv.EntityID={}".format(b_priceID)
                res = self.db.rs(sql)
                if res:
                    b_thick = (res[0][0])
        return band(b_priceID, b_name, b_width, b_thickPan, b_thick, b_cnt)


    def band_b(self, unitpos):
        '''информация о кромке торца B'''
        
        res = self.band_side(unitpos, 7, 1)
        return res


    def band_c(self, unitpos):
        '''информация о кромке торца C'''

        res = self.band_side(unitpos, 3, 1)
        return res


    def band_d(self, unitpos):
        '''информация о кромке торца D'''

        res = self.band_side(unitpos, 1, 1)
        return res


    def band_e(self, unitpos):
        '''информация о кромке торца E'''
        
        res = self.band_side(unitpos, 5, 1)
        return res


    def band_x1(self, unitpos):
        "Кромка вдоль длины одной стороны с учётом текстуры"
        
        pdir = self.dir(unitpos)
        if (45<pdir<=135):
            return self.band_b(unitpos)
        elif (225<pdir<=315):
            return self.band_c(unitpos)
        else:
            return self.band_e(unitpos)


    def band_x2(self, unitpos):
        "Кромка вдоль длины одной стороны с учётом текстуры"
        
        pdir = self.dir(unitpos)
        if (45<pdir<=135):
            return self.band_c(unitpos)
        elif (225<pdir<=315):
            return self.band_b(unitpos)
        else:
            return self.band_d(unitpos)


    def band_y1(self, unitpos):
        "Кромка вдоль длины одной стороны с учётом текстуры"
        
        pdir = self.dir(unitpos)
        if (45<pdir<=135):
            return self.band_e(unitpos)
        elif (225<pdir<=315):
            return self.band_d(unitpos)
        else:
            return self.band_b(unitpos)


    def band_y2(self, unitpos):
        "Кромка вдоль длины одной стороны с учётом текстуры"
        
        pdir = self.dir(unitpos)
        if (45<pdir<=135):
            return self.band_d(unitpos)
        elif (225<pdir<=315):
            return self.band_e(unitpos)
        else:
            return self.band_c(unitpos)


    def decorates(self, unitpos, map=5):
        '''Определяет отделку панели выбранной карты map:
            1 - сторона E (Y+)
            2 - сторона D (Y-)
            3 - сторона C (X+)
            4 - сторона B (X-)
            5 - пласть A (Z+)
            6 - пласть F (Z-)
            7 - угол 1
            8 - угол 2
            9 - угол 3
            10 - угол 4
            11 - дополнение 1
            12 - дополнение 2
            -1 - все стороны
            -2 - все торцы'''

        sql = "SELECT td.materialid, td.typename, tnn.name FROM TDecorates AS td LEFT JOIN TNNomenclature AS tnn ON td.materialid=tnn.id " \
              "WHERE unitpos = {} AND map = {}".format(unitpos, map)
        decor = self.db.rs(sql)
        abr = {'шпон': 'шп.', 'эмаль': 'эм.', 'пластик': 'пл.', 'плёнка (пвх)': 'ПВХ ', 'патина': 'пт.', 'морилка': 'мор.', 'лак': ''}
        rep = ["шп.", "шпон", "Шпон", "эм.", "Эм.", "эмаль", "Эмаль"]
        cntdec = len(decor)
        decorname = ""
        for i in range(cntdec):
            typemat = decor[i][1]
            matname = decor[i][2]
            pref = typemat.lower()
            for j in range(len(rep)):
                matname = matname.replace(rep[j], "")
            if pref in abr.keys():

                decorname += abr[typemat.lower()] + matname
            else:
                decorname += typemat + matname
            if i>=0 and i<(cntdec-1):
                decorname += "+"

        return decorname


    def cnt_drill_pans(self, tpp=None, hingoff=False):
        '''Получить кол-во просверленных деталей. Сверловка с двух сторон считается, как две детали
           hingoff = True - Считать вместе с присадкой под петли
        '''
        filtrtpp = "AND TElems.TopParentPos={0}".format(tpp) if tpp is not None else ""
        # Выбор сверловки всего, кроме стороны F
        sql_exF = "SELECT DISTINCT th.UnitPos FROM THoles AS th LEFT JOIN TElems ON th.UnitPos = TElems.UnitPos " \
                  "WHERE (th.MatrA33<=0 AND (th.UnitPos In (SELECT UnitPos FROM TPanels)) {0})".format(filtrtpp)
        # Выбор отвертий в пласти F
        sql_F = "SELECT DISTINCT th.UnitPos FROM THoles AS th LEFT JOIN TElems ON th.UnitPos = TElems.UnitPos " \
              "WHERE (Abs(th.MatrA13)<0.001 AND Abs(th.MatrA23)<0.001 AND th.MatrA33>0 " \
              "AND th.UnitPos In (SELECT UnitPos FROM TPanels) {0})".format(filtrtpp)
        
        sql_un = " UNION ".join([sql_exF, sql_F])
        sql = "SELECT Count(*) FROM ({})".format(sql_un)
        res = self.db.rs(sql)[0][0]
        
        if hingoff==True:
            return res - self.cnt_pan_hings(tpp=tpp)
        else:
            return res


    def cnt_holes_pan(self, unitpos, hingoff=False):
        '''Получить кол-во отверстий конкретной панели
           hingoff - Считать вместе с присадкой под петли или исключить петли
        '''
        sql = "SELECT Count(*) FROM THoles WHERE THoles.UnitPos={0}".format(unitpos)
        res = self.db.rs(sql)
        if not res:
            return 0
        if hingoff==True:
            return res[0][0] - self.cnt_holes_hings(unitpos=unitpos)
        else:
            return res[0][0]


    def cnt_holes_pan_diam(self, unitpos, d):
        '''Получить кол-во отверстий определённого диаметра конкретной панели'''
        
        sql = "SELECT Count(*) FROM THoles WHERE THoles.UnitPos={0} AND THoles.Diameter={1}".format(unitpos, d)
        res     = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]


    def cnt_holes_hings(self, unitpos=None, tpp=None):
        '''Получить кол-во отверстий под чашку пети у всех панелей, 
           при unitpos=None или конкретной панели, при unitpos=UnitPos панели.
           tpp=TopParentPos - id родительского объекта
        '''
        if unitpos and tpp is None:
            filtrup = "AND th.UnitPos = {0} ".format(unitpos)
        else:
            filtrup = ""
        if tpp:
            filtrtpp = "AND te.TopParentPos = {0}".format(filtrtpp)
        else:
            filtrtpp = ""
        sql = "SELECT Count(th.UnitPos) FROM THoles AS th LEFT JOIN TElems AS te ON th.HolderPos = te.UnitPos WHERE te.FurnType Like '0406%' "\
              "AND th.Diameter > 20 {0}".format(filtrup+filtrtpp)
        res     = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]


    def cnt_hings_x(self, unitpos):
        '''Количество петель вдоль панели'''
        
        sql = "SELECT Count(th.MatrA24) FROM THoles AS th LEFT JOIN TElems AS te ON th.HolderPos = te.UnitPos WHERE te.FurnType Like '0406%' "\
              "AND th.Diameter > 20 AND th.UnitPos = {0} GROUP BY th.MatrA24 HAVING Count(1)>1".format(unitpos)
        res = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]


    def cnt_hings_y(self, unitpos):
        '''Количество петель поперёк панели'''
        
        sql = "SELECT Count(th.MatrA14) FROM THoles AS th LEFT JOIN TElems AS te ON th.HolderPos = te.UnitPos WHERE te.FurnType Like '0406%' "\
              "AND th.Diameter > 20 AND th.UnitPos = {0} GROUP BY th.MatrA14 HAVING Count(1)>1".format(unitpos)
        res = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]


    def cnt_pan_hings(self, tpp=None):
        '''Получить кол-во деталей с присадкой только под петли (без других отв.)
           tpp=TopParentPos - id родительского объекта
        '''
        if tpp:
            filtrtpp = "AND te.TopParentPos = {0}".format(filtrtpp)
        else:
            filtrtpp = ""
        sql = "SELECT Count(*) FROM (SELECT DISTINCT th.UnitPos FROM THoles AS th LEFT JOIN TElems AS te "\
              "ON th.HolderPos = te.UnitPos WHERE (te.FurnType) Like '0406%' {0} AND th.UnitPos not in (SELECT th.UnitPos "\
              "FROM THoles AS th LEFT JOIN TElems AS te ON th.HolderPos = te.UnitPos WHERE (te.FurnType) not Like '0406%' {0}))".format(filtrtpp)
        res = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]

    def cnt_chamfer_pan(self, unitpos):
        '''Определение кол-ва скосов углов на панеле'''

        sql = "SELECT Count(*) FROM TParams WHERE ParamName = 'Cuttype' AND NumValue = 1 AND UnitPos = {0}".format(unitpos)
        res = self.db.rs(sql)
        if not res:
            return 0
        return res[0][0]


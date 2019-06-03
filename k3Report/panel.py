# -*- coding: utf-8 -*-
__author__ = 'Виноградов А.Г. г.Белгород  август 2015'


class Panel:
    def __init__(self, db):
        '''Информация о панели'''
        try:
            self.db = db
        except:
            return None

    def RS(self, sql, typ='n'):
        '''n - ответ число s - ответ строка b - ответ логический'''
        if typ == 'n':
            err = 0
        elif typ == 's':
            err = ''
        elif typ == 'b':
            err = False
        else:
            err = None
        try:
            self.db.cur.execute(sql)
            rows = self.db.cur.fetchall()
            if rows[0][0] is None:
                return err
            return rows[0][0]
        except:
            return err

    def list_panels(self, matid=None, tpp=None):
        '''Получить список панелей по материалу
           matid - id материала из которого сделана панель
           tpp - главный родитель панели
        '''
        filtrid = "AND tnn.ID={}".format(matid)
        if matid is None:
            filtrid = ""
        filtrtpp = " AND te.TopParentPos={}".format(tpp)
        if tpp is None:
            filtrtpp = ""
        sql = "SELECT te.UnitPos FROM TElems AS te INNER JOIN TNNomenclature AS tnn "\
              "ON te.PriceID = tnn.ID WHERE te.FurnType Like '01%' {0}".format(filtrid+filtrtpp)
        res = self.db.recordset(sql)
        id = []
        for i in res:
            id.append(i[0])
        return id

    def xunit(self, unitpos):
        sql = "SELECT xunit FROM TElems WHERE unitpos = {}".format(unitpos)
        return round(self.RS(sql),1)

    def yunit(self, unitpos):
        sql = "SELECT yunit FROM TElems WHERE unitpos = {}".format(unitpos)
        return round(self.RS(sql),1)

    def zunit(self, unitpos):
        sql = "SELECT zunit FROM TElems WHERE unitpos = {}".format(unitpos)
        return round(self.RS(sql),1)

    def cnt(self, unitpos):
        sql = "SELECT count FROM TElems WHERE unitpos = {}".format(unitpos)
        return self.RS(sql)

    def length(self, unitpos):
        sql = "SELECT length FROM TPanels WHERE unitpos = {}".format(unitpos)
        return round(self.RS(sql),1)

    def width(self, unitpos):
        sql = "SELECT width FROM TPanels WHERE unitpos = {}".format(unitpos)
        return round(self.RS(sql),1)

    def planelength(self, unitpos):
        sql = "SELECT planelength FROM TPanels WHERE unitpos = {}".format(unitpos)
        return round(self.RS(sql),1)

    def planewidth(self, unitpos):
        sql = "SELECT planewidth FROM TPanels WHERE unitpos = {}".format(unitpos)
        return round(self.RS(sql),1)

    def thickness(self, unitpos):
        sql = "SELECT thickness FROM TPanels WHERE unitpos = {}".format(unitpos)
        return round(self.RS(sql),1)

    def dir(self, unitpos):
        sql = "SELECT dir FROM TPanels WHERE unitpos = {}".format(unitpos)
        return self.RS(sql)

    def curvepath(self, unitpos):
        sql = "SELECT curvepath FROM TPanels WHERE unitpos = {}".format(unitpos)
        return self.RS(sql, 'b')

    def form(self, unitpos):
        sql = "SELECT formtype FROM TPanels WHERE unitpos = {}".format(unitpos)
        return self.RS(sql)

    def slots(self, unitpos):
        sql = "SELECT iif(Count(panelpos)>0, 1 , 0) FROM TSlots WHERE panelpos = {}".format(unitpos)
        return self.RS(sql, 'b')
    
    def slots_x(self, unitpos):
        '''Количество пропилов вдоль длины'''
        sql = "SELECT Count(PanelPos) FROM TSlots WHERE ROUND(BegY)=ROUND(EndY) AND PanelPos = {0}".format(unitpos)
        res = self.RS(sql)
        return res

    def slots_y(self, unitpos):
        '''Количество пропилов поперёк длины'''
        sql = "SELECT Count(PanelPos) FROM TSlots WHERE ROUND(BegX)=ROUND(EndX) AND PanelPos = {0}".format(unitpos)
        res = self.RS(sql)
        return res

    def slots_x_len(self, unitpos):
        '''Суммарная длина пропилов вдоль длины'''
        sql = "SELECT ROUND(SUM(ABS(EndX-BegX))/1000,2) FROM TSlots WHERE ROUND(BegY)=ROUND(EndY) AND PanelPos = {0}".format(unitpos)
        res = self.RS(sql)
        return res

    def slots_y_len(self, unitpos):
        '''Суммарная длина пропилов поперёк длины'''
        sql = "SELECT ROUND(SUM(ABS(EndY-BegY))/1000,2) FROM TSlots WHERE ROUND(BegX)=ROUND(EndX) AND PanelPos = {0}".format(unitpos)
        res = self.RS(sql)
        return res

    def butts(self, unitpos):
        sql = "SELECT iif(Count(unitpos)>0, 1 , 0) FROM TButts WHERE unitpos = {}".format(unitpos)
        return self.RS(sql, 'b')

    def frezerovka(self, unitpos):
        sql = "SELECT attrstring FROM TAttributes WHERE name='frezerovka' AND unitpos = {}".format(unitpos)
        return self.RS(sql, 's')

    def priceid(self, unitpos):
        sql = "SELECT priceid FROM TElems WHERE unitpos = {}".format(unitpos)
        return self.RS(sql)

    def name(self, unitpos):
        sql = "SELECT name FROM TElems WHERE unitpos = {}".format(unitpos)
        return self.RS(sql, 's')

    def data(self, unitpos):
        sql = "SELECT data FROM TElems WHERE unitpos = {}".format(unitpos)
        return self.RS(sql, 's')

    def cmpos(self, unitpos):
        sql = "SELECT commonpos FROM TElems WHERE unitpos = {}".format(unitpos)
        return self.RS(sql)

    def ppos(self, unitpos):
        sql = "SELECT parentpos FROM TElems WHERE unitpos = {}".format(unitpos)
        return self.RS(sql)

    def tpos(self, unitpos):
        sql = "SELECT topparentpos FROM TElems WHERE unitpos = {}".format(unitpos)
        return self.RS(sql)

    def furntype(self, unitpos):
        sql = "SELECT furntype FROM TElems WHERE unitpos = {}".format(unitpos)
        return self.RS(sql, 's')

    def sumcost(self, unitpos):
        sql = "SELECT sumcost FROM TElems WHERE unitpos = {}".format(unitpos)
        return round(self.RS(sql),2)

    def svw(self, unitpos):
        '''s - Площадь кв.м, v - объём куб.м, w - вес кг, density - плотность'''
        p = Panel(self)
        id = p.priceid(unitpos)
        sql = "SELECT tnpv.DValue FROM TNProperties AS tnp LEFT JOIN TNPropertyValues AS tnpv "\
              "ON tnp.ID = tnpv.PropertyID WHERE (((tnp.Ident)='density') AND ((tnpv.EntityID)={}))".format(id)
        density = int(self.RS(sql))
        s = p.xunit(unitpos)*p.yunit(unitpos)/1000000
        v = s*p.zunit(unitpos)/1000
        w = v*density
        return round(s,2), round(v,4), round(w,2)

    def radius(self, unitpos):
        '''информация о радиусе гнутой панели и оси гиба'''
        radius = 0
        bend = 0            # ось гиба: 1 - OX 2 - OY
        p = Panel(self)
        form = p.form(unitpos)  # узнаём форму панели

        if form == 1:           # дуга по хорде
            sql = "SELECT abs(tpm1.NumValue) AS Caving, tpm2.NumValue AS Chord, tpm3.NumValue AS Bend "\
                  "FROM TParams AS tpm1, TParams AS tpm2, TParams AS tpm3 "\
                  "WHERE (((tpm1.UnitPos)={0}) AND ((tpm1.ParamName)='ArcChord.Caving') AND ((tpm2.UnitPos)={0}) AND "\
                  "((tpm2.ParamName)='ArcChord.Chord') AND ((tpm3.UnitPos)={0}) AND ((tpm3.ParamName)='BendAxis'))".format(unitpos)
            self.db.cur.execute(sql)
            rows = self.db.cur.fetchall()[0]
            caving = rows[0]
            chord = rows[1]
            bend = rows[2]
            alfa = 2*math.atan((2*caving)/chord)
            L = chord*(alfa/math.sin(alfa))
            D = L/alfa
            radius = D/2

        if form == 2:           # Два отрезка и дуга
            sql = "SELECT tpm1.NumValue, tpm2.NumValue AS Bend FROM TParams AS tpm1, TParams AS tpm2 "\
                  "WHERE (tpm1.UnitPos={0} AND tpm1.ParamName='LinesArc.R') AND (tpm2.UnitPos={0} "\
                  "AND tpm2.ParamName='BendAxis')".format(unitpos)
            self.db.cur.execute(sql)
            rows = self.db.cur.fetchall()[0]
            radius = rows[0]
            bend = rows[1]

        return round(radius,1), int(bend)

    def band_side(self, unitpos, IDLine, IDPoly=1):
        '''информация о кромке торца'''
        b_priceID = 0       # id кромки в номенклатуре
        b_name = ''         # название кромки
        b_width = 0         # ширина кромки
        b_thickPan = 0      # толщина панели
        b_thick = 0         # толщина кромки
        b_cnt = 0
        keys = ('ID', 'Name', 'Width', 'ThickPan', 'ThickBand', 'Count')
        # определим тип панели
        # PanPolyType: 1 - полилиния 2 - прямоугольная 3 - четырёхугольная
        sql = "SELECT tpm.NumValue FROM TParams AS tpm WHERE (tpm.unitpos = {} AND tpm.ParamName='PanPolyType' AND tpm.HoldTable='TPanels')".format(unitpos)
        PolyType = int(self.RS(sql))
        if (PolyType == 2) or (PolyType == 3):
            sql = "SELECT bu.BandUnitPos, tn.ID, tn.Name, Count FROM (SELECT idp.NumValue AS IDPoly, idl.NumValue AS IDLine, idb.NumValue AS BandUnitPos "\
                  "FROM (TParams idp LEFT OUTER JOIN TParams idl ON idp.UnitPos=idl.UnitPos AND idp.HoldTable=idl.HoldTable AND idp.Hold1=idl.Hold1 "\
                  "AND idp.Hold3=idl.Hold3) LEFT OUTER JOIN TParams idb ON idp.UnitPos=idb.UnitPos AND idp.HoldTable=idb.HoldTable AND idp.Hold1=idb.Hold1 "\
                  "AND idp.Hold3=idb.Hold3 WHERE idp.UnitPos={0} AND idp.HoldTable='TPaths' AND idp.Hold1=1 AND idp.ParamName='IDPoly' AND idl.ParamName='IDLine' "\
                  "AND idb.ParamName='BandUnitPos') bu, TNNomenclature tn, TElems te WHERE bu.BandUnitPos=te.UnitPos AND te.PriceID=tn.ID AND IDPoly={1} AND IDLine={2}".format(unitpos, IDPoly, IDLine)
            res = self.db.recordset(sql)
            if res:
                bandID = (res[0][0])
                b_priceID = (res[0][1])
                b_name = (res[0][2])
                b_cnt = (res[0][3])
                sql = "SELECT Width FROM TBands WHERE UnitPos={}".format(bandID)
                res = self.db.recordset(sql)
                if res:
                    b_thickPan = (res[0][0])
                sql = "SELECT tnpv.DValue FROM TNPropertyValues AS tnpv WHERE tnpv.PropertyID=21 AND tnpv.EntityID={}".format(b_priceID)
                res = self.db.recordset(sql)
                if res:
                    b_width = (res[0][0])
                sql = "SELECT tnpv.DValue FROM TNPropertyValues AS tnpv WHERE tnpv.PropertyID=10 AND tnpv.EntityID={}".format(b_priceID)
                res = self.db.recordset(sql)
                if res:
                    b_thick = (res[0][0])            
        return dict(zip(keys,(b_priceID, b_name, b_width, b_thickPan, b_thick, b_cnt)))

    def band_b(self, unitpos):
        '''информация о кромке торца B'''
        res = (self.band_side(unitpos, 7, 1))
        return res

    def band_c(self, unitpos):
        '''информация о кромке торца C'''
        res = (self.band_side(unitpos, 3, 1))
        return res

    def band_d(self, unitpos):
        '''информация о кромке торца D'''
        res = (self.band_side(unitpos, 1, 1))
        return res

    def band_e(self, unitpos):
        '''информация о кромке торца E'''
        res = (self.band_side(unitpos, 5, 1))
        return res

    def decorates(self, unitpos, map):
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
        decor = db.recordset(sql)
        abr = {'шпон':'шп.', 'эмаль':'эм.', 'пластик':'пл.', 'плёнка (пвх)':'ПВХ ', 'патина':'пт.', 'морилка':'мор.', 'лак':''}
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
           hingoff - Считать вместе с присадкой под петли или исключить петли
        '''
        filtrtpp = "WHERE te.TopParentPos={}".format(tpp) if tpp is not None else ""
        sql = "SELECT Sum(te.Count) AS cnt FROM (SELECT DISTINCT th.UnitPos FROM THoles AS th WHERE ABS(MatrA33)<0.001 AND "\
              "(ABS(MatrA13)<0.001 OR ABS (MatrA23)<0.001) AND UnitPos IN (SELECT UnitPos FROM TPanels) "\
              "UNION "\
              "SELECT DISTINCT th.UnitPos FROM THoles AS th WHERE ABS(MatrA13)<0.001 AND ABS(MatrA23)<0.001 AND UnitPos IN (SELECT UnitPos FROM TPanels) AND MatrA33<0 "\
              "UNION "\
              "SELECT DISTINCT th.UnitPos FROM THoles AS th WHERE ABS(MatrA13)<0.001 AND ABS(MatrA23)<0.001 AND UnitPos IN (SELECT UnitPos FROM TPanels) AND MatrA33>0)  AS holes "\
              "LEFT JOIN TElems AS te ON holes.UnitPos = te.UnitPos {0}".format(filtrtpp)
        res = self.db.recordset(sql)
        if hingoff==True:
            return res[0][0] - self.cnt_pan_hings(tpp=tpp)
        else:
            return res[0][0]
    
    def cnt_holes_pan(self, up, hingoff=False):
        '''Получить кол-во отверстий конкретной панели
           hingoff - Считать вместе с присадкой под петли или исключить петли
        '''
        sql = "SELECT Count(*) FROM THoles WHERE THoles.UnitPos={0}".format(up)
        res = self.db.recordset(sql)
        if hingoff==True:
            return res[0][0] - self.cnt_holes_hings(up=up)
        else:
            return res[0][0]
    
    def cnt_holes_pan_diam(self, up, d):
        '''Получить кол-во отверстий определённого диаметра конкретной панели'''
        sql = "SELECT Count(*) FROM THoles WHERE THoles.UnitPos={0} AND THoles.Diameter={1}".format(up, d)
        res     = self.db.recordset(sql)
        return res[0][0]

    def cnt_holes_hings(self, up=None, tpp=None):
        '''Получить кол-во отверстий под чашку пети у всех панелей, 
           при up=None или конкретной панели, при up=UnitPos панели.
           tpp=TopParentPos - id родительского объекта
        '''
        if up and tpp is None:
            filtrup = "AND th.UnitPos = {0}".format(up)
        else:
            filtrup = ""
        if tpp:
            filtrtpp = "AND te.TopParentPos = {0}".format(filtrtpp)
        else:
            filtrtpp = ""
        sql = "SELECT Count(th.UnitPos) FROM THoles AS th LEFT JOIN TElems AS te ON th.HolderPos = te.UnitPos WHERE te.FurnType Like '0406%' "\
              "AND th.Diameter > 20 {0}".format(filtrup+filtrtpp)
        return self.RS(sql)
    
    def cnt_hings_x(self, up):
        '''Количество петель вдоль панели'''
        sql = "SELECT Count(th.MatrA24) FROM THoles AS th LEFT JOIN TElems AS te ON th.HolderPos = te.UnitPos WHERE te.FurnType Like '0406%' "\
              "AND th.Diameter > 20 AND th.UnitPos = {0} GROUP BY th.MatrA24 HAVING Count(1)>1".format(up)
        res = self.RS(sql)
        return res
    
    def cnt_hings_y(self, up):
        '''Количество петель поперёк панели'''
        sql = "SELECT Count(th.MatrA14) FROM THoles AS th LEFT JOIN TElems AS te ON th.HolderPos = te.UnitPos WHERE te.FurnType Like '0406%' "\
              "AND th.Diameter > 20 AND th.UnitPos = {0} GROUP BY th.MatrA14 HAVING Count(1)>1".format(up)
        res = self.RS(sql)
        return res
    
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
        res = self.RS(sql)
        return res

    def cnt_chamfer_pan(self, up):
        '''Определение кол-ва скосов углов на панеле'''
        sql = "SELECT Count(*) FROM TParams WHERE ParamName = 'Cuttype' AND NumValue = 1 AND UnitPos = {0}".format(up)
        res = self.RS(sql)
        return res
__author__ = "Виноградов А.Г. г.Белгород 2018"

# Выводит отчёт деталировки на распил

import k3
from pyRep.MReports import *
import itertools
import operator
import os

import time

LISTKEY = (
    "num",
    "name",
    "length",
    "width",
    "band_x04",
    "band_y04",
    "band_x2",
    "band_y2",
    "chamfer",
    "slots_x_len",
    "slots_y_len",
    "holes",
    "cnt_hings_x",
    "cnt_hings_y",
    "gluing",
    "color_band",
)


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

    return [
        {"_key": i, "data": list(map(_groupid_drop, grp))}
        for i, grp in itertools.groupby(
            lst_dpars, operator.itemgetter(*map(lambda nm: nm, listkey))
        )
    ]


def list_pan(matid, tpp=None):
    """Получаем список деталей панелей"""
    idPan = pn.list_panels(matid, tpp)
    listPan = []
    listPan2 = []
    for i in idPan:
        dic = {}
        cb = ""
        if pn.form(i) == 0:
            telems = bs.telems(i)
            dic["UnitPos"] = i
            dic["num"] = telems["CommonPos"]
            dic["name"] = telems["Name"]
            dic["length"] = pn.planelength(i)
            dic["width"] = pn.planewidth(i)
            pdir = pn.dir(i)
            dic["cnt"] = telems["Count"]

            bb = pn.band_b(i)
            bc = pn.band_c(i)
            bd = pn.band_d(i)
            be = pn.band_e(i)
            artbb = nm.property_name(bb["ID"], "Article")
            artbc = nm.property_name(bc["ID"], "Article")
            artbd = nm.property_name(bd["ID"], "Article")
            artbe = nm.property_name(be["ID"], "Article")
            bbna = bb["Name"] + (" " + artbb if artbb else "")
            bcna = bc["Name"] + (" " + artbc if artbc else "")
            bdna = bd["Name"] + (" " + artbd if artbd else "")
            bena = be["Name"] + (" " + artbe if artbe else "")
            cb = " ".join(set([bbna, bcna, bdna, bena]))

            dic["band_x04"] = int(0 < be["ThickBand"] < 2) + int(
                0 < bd["ThickBand"] < 2
            )
            dic["band_y04"] = int(0 < bb["ThickBand"] < 2) + int(
                0 < bc["ThickBand"] < 2
            )
            dic["band_x2"] = int(0 < be["ThickBand"] == 2) + int(
                0 < bd["ThickBand"] == 2
            )
            dic["band_y2"] = int(0 < bb["ThickBand"] == 2) + int(
                0 < bc["ThickBand"] == 2
            )
            dic["length"] += fuga * (
                dic["band_y04"] + dic["band_y2"]
            )  # добавляем к размеру прифуговку
            dic["width"] += fuga * (
                dic["band_x04"] + dic["band_x2"]
            )  # добавляем к размеру прифуговку
            dic["chamfer"] = pn.cnt_chamfer_pan(i)
            dic["slots_x_len"] = pn.slots_x_len(
                i
            )  # Суммарная длина пропилов вдоль панели
            dic["slots_y_len"] = pn.slots_y_len(
                i
            )  # Суммарная длина пропилов поперёк панели
            dic["holes"] = pn.cnt_holes_pan(i, hingoff=True)  # любые отверстия
            dic["cnt_hings_x"] = pn.cnt_hings_x(i)  # Количество петель вдоль панели
            dic["cnt_hings_y"] = pn.cnt_hings_y(i)  # Количество петель поперёк панели
            dic["gluing"] = nm.property_name(matid, "gluing")  # склейка
            dic["color_band"] = cb
            if (45 < pdir <= 135) or (225 < pdir <= 315):
                dic["band_x04"], dic["band_y04"] = dic["band_y04"], dic["band_x04"]
                dic["band_x2"], dic["band_y2"] = dic["band_y2"], dic["band_x2"]
                dic["slots_x_len"], dic["slots_y_len"] = (
                    dic["slots_y_len"],
                    dic["slots_x_len"],
                )
                dic["cnt_hings_x"], dic["cnt_hings_y"] = (
                    dic["cnt_hings_y"],
                    dic["cnt_hings_x"],
                )
            listPan.append(dic)

    for i in group_list_panels(listPan):
        dic = dict(zip(LISTKEY, i["_key"]))
        cnt = 0
        for j in i["data"]:
            cnt += j["cnt"]
        dic.update({"cnt": cnt})
        listPan2.append(dic)
    listPan2.sort(key=operator.itemgetter("length", "width"), reverse=True)
    return listPan2


def newsheet(name):
    "Создаём новый лист с именем"
    xl.sheetorient = xl.xlconst["xlLandscape"]
    xl.rightmargin = 0.6
    xl.leftmargin = 0.6
    xl.bottommargin = 1.6
    xl.topmargin = 1.6
    xl.centerhorizontally = 0
    xl.displayzeros = False
    xl.new_sheet(name)
    return


def rep_pan(name, mat, tpp=None):
    "Вывод деталей панелей"
    fnum = "0,0"  # числовой формат
    frub = '_-* # 0_ р_-;-* # 0_ р_-;_-* "-"?? р_-;_-@_-'  # денежный формат
    frub2 = '_-* # ##0,00_ р_-;-* # ##0,00_ р_-;_-* "-"?? р_-;_-@_-'  # денежный формат
    fgen = ""  # общий формат
    txt = "@"  # текстовый формат
    newsheet(name)
    row = 1
    rw1 = 1
    if fuga > 0:
        xl.putval(row, 1, "Прифуговка на кромку {}мм".format(fuga))
        row += 1
        rw1 += 1
    # 'Выводим таблицу листовых материалов'

    row += 1
    for imat in mat:
        for l in list_pan(imat["ID"], tpp):
            val = (
                l["name"],
                l["num"],
                imat["Thickness"],
                l["length"],
                l["width"],
                l["cnt"],
                l["band_x04"],
                l["band_y04"],
                l["band_x2"],
                l["band_y2"],
                "",
                "",
                "",
                l["chamfer"],
                "",
                "",
                l["slots_x_len"],
                l["slots_y_len"],
                l["cnt_hings_x"],
                l["cnt_hings_y"],
                l["holes"],
                l["gluing"],
                l["color_band"],
                imat["Name"],
            )
            xl.putval(row, 1, val)
            xl.txtformat(row, 17, nf=[fnum, fnum])
            xl.txtformat(row, 4, halign="r", nf=["0", "0"])
            row += 1
    draw_table(rw1, row)


def draw_table(rw1, rw2):
    "Форматируем таблицу"
    head = [
        "Наименование",
        "№ п/п",
        "Толщина",
        "Высота,  мм",
        "Ширина, мм",
        "шт",
        "0,4 мм",
        "",
        "2  мм",
        "",
        "Отметка",
        "04 кр",
        "2 кр",
        "Угол, шт",
        "Радиус",
        "Фрезер",
        "Паз, п/м",
        "",
        "Петли, шт",
        "",
        "Присадка, шт",
        "Склейка",
        "Цвет кромки",
        "Материал",
    ]
    sp_tab = [
        19,
        2.4,
        6.7,
        8,
        8,
        6.6,
        2,
        2,
        2,
        2,
        1.6,
        3,
        3,
        3,
        3,
        3,
        3,
        3,
        3,
        3,
        3,
        3,
        7,
        25,
    ]
    xl.columnsize(rw1, sp_tab)
    xl.rowsize(rw1, 52)
    xl.header(rw1, 1, head, tc=0, pat=0)
    xl.mergecells(rw1, 7, 2)
    xl.mergecells(rw1, 9, 2)
    xl.mergecells(rw1, 17, 2)
    xl.mergecells(rw1, 19, 2)
    xl.txtformat(
        rw1,
        1,
        halign="c",
        valign="c",
        wrap="t",
        ort=[
            0,
            0,
            90,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            90,
            90,
            90,
            90,
            90,
            90,
            0,
            0,
            0,
            0,
            90,
            90,
            0,
            0,
        ],
        fsz=[
            12,
            8,
            8,
            12,
            12,
            12,
            12,
            12,
            12,
            12,
            8,
            12,
            12,
            12,
            12,
            12,
            12,
            12,
            12,
            12,
            8,
            8,
            12,
            12,
        ],
    )
    gh = rw2 - rw1
    xl.gridset(tc=2)
    xl.grid(rw1, 1, 24, gh, "lrudvh")
    xl.gridset(wt="xlMedium", tc=2)
    xl.grid(rw1, 4, 2, gh, "lr")
    xl.grid(rw1, 7, 2, gh, "lr")
    xl.grid(rw1, 9, 2, gh, "lr")
    xl.grid(rw1, 17, 2, gh, "lr")
    xl.grid(rw1, 19, 2, gh, "lr")
    # форматируем условно
    xl.wb.ActiveSheet.Range("L:V").FormatConditions.Add(1, 5, "=0")
    gr = xl.wb.ActiveSheet.Range("L:V").FormatConditions(1).Interior
    gr.PatternColorIndex = -4105
    gr.ThemeColor = 7
    gr.TintAndShade = 0.799981688894314
    xl.wb.ActiveSheet.Range("D:E").FormatConditions.Add(1, 5, "=2800")
    gr = xl.wb.ActiveSheet.Range("D:E").FormatConditions(1).Interior
    gr.PatternColorIndex = -4105
    gr.ThemeColor = 6
    gr.TintAndShade = 0.399945066682943


def start(tpp=None):
    # 'Формируем данные'
    matid = nm.matbyuid(2, tpp)
    listmat = []
    listDSP = []
    listMDF = []
    listDVP = []
    listGLS = []
    if matid:
        for i in matid:
            dic = {}
            prop = nm.properties(i)
            dic["ID"] = i
            dic["Name"] = prop.get("Name")
            dic["Thickness"] = int(prop.get("Thickness"))
            dic["MatTypeID"] = int(prop.get("MatTypeID"))
            if dic["MatTypeID"] == 128:
                listDSP.append(dic)
                listDSP.sort(key=lambda x: int(x["Thickness"]), reverse=True)
            elif dic["MatTypeID"] == 64:
                listMDF.append(dic)
                listMDF.sort(key=lambda x: int(x["Thickness"]), reverse=True)
            elif dic["MatTypeID"] == 37:
                listDVP.append(dic)
                listDVP.sort(key=lambda x: int(x["Thickness"]), reverse=True)
            elif dic["MatTypeID"] in [48, 99]:
                listGLS.append(dic)
                listGLS.sort(key=lambda x: int(x["Thickness"]), reverse=True)
            else:
                listmat.append(dic)
                listmat.sort(
                    key=lambda x: [int(x["MatTypeID"]), int(x["Thickness"])],
                    reverse=True,
                )

        if listDSP:
            rep_pan("ДСП", listDSP, tpp)

        if listMDF:
            rep_pan("МДФ", listMDF, tpp)

        if listDVP:
            rep_pan("ДВП", listDVP, tpp)

        if listGLS:
            rep_pan("Стекло", listGLS, tpp)

        if listmat:
            rep_pan("Прочее", listmat, tpp)


if __name__ == "__main__":

    # file = r'd:\PKMProjects74\17\17.mdb'
    # projreppath = r'd:\PKMProjects74\17'
    # fuga = 0
    starttime = time.time()
    params = k3.getpar()
    file = params[0]
    projreppath = params[1]
    fuga = params[3]

    db = DB()
    tmp = db.connect(file)  # Подключаемся к базе выгрузки
    if tmp == "NoFile":
        raise Exception("Ошибка подключения к базе данных")
    try:
        xl = ExcelDoc()
        xl.Excel.Visible = 0
        xl.Excel.Application.ScreenUpdating = False
        nm = Nomenclature(db)
        bs = Base(db)
        pn = Panel(db)
        start()
        xl.save(os.path.join(projreppath, "Деталировка"))
        xl.Excel.Visible = 1
        xl.Excel.Application.ScreenUpdating = True
        params[2].value = 1
    except:
        raise Exception("Произошла ошибка во время создания отчёта")
    db.disconnect()
    endtime = time.time()
    print(endtime - starttime)

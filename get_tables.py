import math
import k3r
from collections import namedtuple


class Specific:
    """Получение спецификации"""

    def __init__(self, db):
        self.bs = k3r.base.Base(db)
        self.ln = k3r.long.Long(db)
        self.nm = k3r.nomenclature.Nomenclature(db)
        self.pn = k3r.panel.Panel(db)
        self.pr = k3r.prof.Profile(db)
        self.bs = k3r.base.Base(db)

    def t_sheets(self, tpp=None):
        """Таблица листовыйх материалов
        Входные данные:
            tpp - int TopParentPos id объекта (шкафа)
        Вывод:
           priceid, name, mattypeid, article, unitsname, price,
           wastecoeff, pricecoeff, density, thickness, sqm
        """
        keys = ('priceid', 'name', 'mattypeid', 'article', 'unitsname', 'price',
                'wastecoeff', 'pricecoeff', 'density', 'thickness', 'sqm')
        mat_id = self.nm.mat_by_uid(2, tpp)
        sh = []
        for i in mat_id:
            Sheets = namedtuple('Sheets', keys)
            p = self.nm.properties(i)
            density = getattr(p, 'density', 0)
            sqm = self.nm.sqm(i, tpp)
            sheet = Sheets(p.priceid, p.name, p.mattypeid, p.article, p.unitsname, p.price,
                           p.wastecoeff, p.pricecoeff, density, p.thickness, sqm)
            sh.append(sheet)
        sh.sort(key=lambda x: [int(x.mattypeid), int(x.thickness)], reverse=True)
        return sh

    def t_acc(self, tpp=None, uid=None):
        """Таблица комлпектующих. Не включает погонажные изделия типа сеток
        Возвращает:
            priceid - из номенклатуры
            cnt - количество
            ... - набор свойств присущих номенклатурному материалу
        """
        acc_list = self.nm.acc_by_uid(uid, tpp)
        acc = []
        for i in acc_list:
            prop = self.nm.properties(i.priceid)
            obj = k3r.utils.tuple_append(prop, {'cnt': i.cnt}, 'Acc')
            acc.append(obj)
        return acc

    def t_acc_long(self, tpp=None):
        """Таблица погонажных комлпектующих.
        Возвращает:
            priceid - из номенклатуры
            len - длина в метрах
            cnt - количество
            ... - набор свойств присущих номенклатурному материалу
        """
        acc_list = self.nm.acc_long(tpp)
        acc = []
        for i in acc_list:
            prop = self.nm.properties(i.priceid)
            obj = k3r.utils.tuple_append(prop, {'len': i.len, 'cnt': i.cnt}, 'Acc')
            acc.append(obj)
        return acc

    def t_bands(self, add=0, tpp=None):
        """Таблица кромок
        Входные данные:
           add - добавочная длина кромки в мм на торец для отходов
           tpp - ID хозяина кромки
       Выходные данные:
           len - длина
           thick - толщина торца
           ... - набор свойств присущих номенклатурному материалу
        """
        bands = self.nm.bands(add, tpp)
        t_bands = []
        for i in bands:
            prop = self.nm.properties(i.priceid)
            obj = k3r.utils.tuple_append(prop, {'len': i.len, 'thick': i.thick})
            t_bands.append(obj)
        return t_bands

    def t_profiles(self, tpp=None):
        """Таблица кусков профилей
        Возвращает:
            priceid - из номенклатуры
            len - длина в метрах
            cnt - количество
            formtype - тип формы профиля
            ... - набор свойств присущих номенклатурному материалу
        """
        pr_list = self.pr.profiles(tpp)
        prof = []
        for i in pr_list:
            prop = self.nm.properties(i.priceid)
            obj = k3r.utils.tuple_append(prop, {'len': i.len, 'formtype': i.formtype, 'cnt': i.cnt}, 'Prof')
            prof.append(obj)
        return prof

    def t_total_prof(self, tpp=None):
        """Таблица профилей
        Возвращает:
            priceid - из номенклатуры
            len - длина в метрах с учётом кратности нарезки
            net_len - длина в местрах в чистом виде
            ... - набор свойств присущих номенклатурному материалу
        """
        pr_list = self.pr.total(tpp)
        prof = []
        for i in pr_list:
            prop = self.nm.properties(i.priceid)
            stepcut = getattr(prop, 'stepcut', 1)
            len = math.ceil(i.len / stepcut) * stepcut
            obj = k3r.utils.tuple_append(prop, {'len': len, 'net_len': i.len}, 'Prof')
            prof.append(obj)
        return prof

    def t_longs(self, tpp=None):
        """Таблица длиномеров
        Вывод: 'type', 'priceid', 'length', 'width', 'height', 'cnt', 'form
        Типы длиномеров:
            0	Столешница
            1	Карниз
            2	Стеновая панель
            3	Водоотбойник
            4	Профиль карниза
            5	Цоколь
            6	Нижний профиль
            7	Балюстрада
        """
        return self.ln.long_list(tpp)


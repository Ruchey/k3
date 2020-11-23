from collections import OrderedDict, namedtuple
from itertools import groupby


def norm_key_prop(lst):
    """В ключах убирает первые цифровые символы, пробел внутри слов удаляет
    Входные данные:
    lst = [('свойство', значение), ('свойство', значение)]
    """

    new_list = []
    for tup in lst:
        key = tup[0].replace(" ", "").lower().lstrip("0123456789.- ")
        new_list.append((key, tup[1]))
    return new_list


def float_int(num):
    """Убираем нули после точки, если число целое"""

    if type(num) == int:
        return num

    if num.is_integer():
        return int(num)
    else:
        return round(num, 1)


def sum_by_key(iterable, key, sum_field):
    """Группирует список по полю key и суммирует по полю sum_field
    iterable - список именованных кортежей
    key - строка, название общего поля для группировки
    sum_field - строка, название поля содержащие кол-во, которое будет суммироваться
    было:
        Pans(cpos=2, name='Стойка правая', length=584, width=400, cnt=1),
        Pans(cpos=2, name='Стойка левая', length=584, width=400, cnt=1),
    стало:
        Pans(cpos=2, name='Стойка правая', length=584, width=400, cnt=2),
    """

    new_list = []
    keys = list(getattr(x, key) for x in iterable)
    dist_keys = list(OrderedDict.fromkeys(keys))
    for i in dist_keys:
        items = list(filter(lambda x: getattr(x, key) == i, iterable))
        cnt = sum(getattr(x, sum_field, 0) for x in items)
        new_list.append(items[0]._replace(**{sum_field: cnt}))
    return new_list


def group_by_keys(iterable, keys, sum_field=None, stick_field=None):
    """Группирует список по полю списку ключей key и суммирует по полю sum_field.
    Дополнительное поле stick_field склеивает значения, например по полю cpos - номер детали.
    На выходе получим cpos = '12;34;2'
    Если поле sum_field не задано, на выходе получим список групп (списков) по схожим заданным ключам.
    Входные данные:
        iterable -- obj список именованных кортежей
        key -- list названия полей для группировки
        sum_field -- str название поля содержащие кол-во, которое будет суммироваться
        stick_field -- str название поля для склейки их значений
        было:
            Pans(cpos=2, name='Стойка правая', length=584, width=400, cnt=1),
            Pans(cpos=3, name='Стойка левая', length=584, width=400, cnt=1),
        стало:
            Pans(cpos='2;3', name='Стойка правая', length=584, width=400, cnt=2),
    """
    dist_keys = []
    new_list = []
    for i in iterable:
        key_set = []
        for j in keys:
            key_set.append(getattr(i, j))
            dist_keys.append(key_set)
            dist_keys.sort()
            dist_keys = [el for el, _ in groupby(dist_keys)]
    for vals in dist_keys:
        kv = list(zip(keys, vals))
        items = [i for i in iterable if all(getattr(i, k) == v for k, v in kv)]
        if sum_field:
            cnt = sum(getattr(x, sum_field) for x in items)
            if stick_field:
                stick_field_list = list(str(getattr(x, stick_field)) for x in items)
                stick = ";".join(OrderedDict.fromkeys(stick_field_list))
                new_list.append(
                    items[0]._replace(**{sum_field: cnt, stick_field: stick})
                )
            else:
                new_list.append(items[0]._replace(**{sum_field: cnt}))
        else:
            new_list.append(items)
    return new_list


def get_tree_parents(unitpos, table):
    """Возвращает список связанных родителей"""
    tree = []
    el = [i for i in table if i[0] == unitpos]
    if el:
        tree.extend(el)
        tree.extend(get_tree_parents(el[0][1], table))
        return tree
    else:
        return tree


def tuple_append(tup, dic, name="NTuple"):
    """Добавление в именованный кортеж данных из словаря"""
    tmp = tup._asdict()
    tmp.update(dic)
    NTuple = namedtuple(name, dict(tmp).keys())
    return NTuple(**dict(tmp))

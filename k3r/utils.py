# -*- coding: utf-8 -*-
# Вспомогательные функции

from collections import OrderedDict

__author__ = 'Виноградов А.Г. г.Белгород июнь 2019'


def norm_key_prop(lst):
    """В ключах убирает первые цифровые символы, пробел внутри слов удаляет
    Входные данные:
    lst = [('свойство', значение), ('свойство', значение)]
    """

    new_list = []
    for tup in lst:
        key = tup[0].replace(" ", "").lower().lstrip('0123456789.- ')
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


def group_by_key(iterable, key, sum_field):
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

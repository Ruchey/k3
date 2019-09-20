# -*- coding: utf-8 -*-
# Вспомогательные функции

from collections import OrderedDict

__author__ = 'Виноградов А.Г. г.Белгород июнь 2019'


def normkeyprop(lst):
    """В ключах убирает первые цифровые символы, пробел внутри слов удаляет
    Входные данные:
    lst = [('свойство', значение), ('свойство', значение)]"""

    newlist = []
    for tup in lst:
        key = tup[0].replace(" ","").lower().lstrip('0123456789.- ')
        newlist.append((key, tup[1]))
    return newlist


def float_int(num):
    'Убираем нули после точки, если число целое'

    if type(num) == int:
        return num

    if num.is_integer():
        return int(num)
    else:
        return round(num, 1)


def groupbykey(iterable, key, sumfield):
    """Группирует список по полю key и суммирует по полю sumfield
    iterable - список именованных кортежей
    key - строка, название общего поля для группировки
    sumfield - строка, название поля содержащие кол-во, которое будет суммироваться
    """
    
    newlist = []
    keys = list(getattr(x, key) for x in iterable)
    distkeys = list(OrderedDict.fromkeys(keys))
    for i in distkeys:
        items = list(filter(lambda x: getattr(x, key)==i, iterable))
        cnt = sum(getattr(x, sumfield) for x in items)
        newlist.append(items[0]._replace(**{sumfield: cnt}))
    return newlist
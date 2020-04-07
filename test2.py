# -*- coding: utf-8 -*-

from collections import OrderedDict, namedtuple
from itertools import groupby

import pprint


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
        cnt = sum(getattr(x, sum_field) for x in items)
        new_list.append(items[0]._replace(**{sum_field: cnt}))
    return new_list


def group_by_keys(iterable, keys, sum_field, stick_field=None):
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
        cnt = sum(getattr(x, sum_field) for x in items)
        if stick_field:
            stick_field_list = list(str(getattr(x, stick_field)) for x in items)
            stick = ';'.join(OrderedDict.fromkeys(stick_field_list))
            new_list.append(items[0]._replace(**{sum_field: cnt, stick_field: stick}))
        else:
            new_list.append(items[0]._replace(**{sum_field: cnt}))
    return new_list



keys = ('cpos', 'name', 'length', 'width', 'cnt')
Pans = namedtuple('Pans', keys)
lp = [Pans(1, 'one', 300, 200, 1),
      Pans(2, 'tow', 200, 200, 2),
      Pans(2, 'three', 100, 200, 3),
      Pans(3, 'for', 200, 200, 1),
      Pans(3, 'wer', 200, 200, 2),
      Pans(3, 'fgt', 200, 200, 1)
      ]

pprint.pprint(group_by_keys(lp, ['width', 'length'], 'cnt', 'cpos'))

# -*- coding: utf-8 -*-

from collections import OrderedDict, namedtuple
from itertools import groupby

import pprint


def group_by_keys(iterable, keys):
    """Группировка в списки объектов по заданным ключам"""
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
        new_list.append(items)
    return new_list


kes = ("cpos", "name", "length", "width", "cnt")
Pans = namedtuple("Pans", kes)
lp = [
    Pans(1, "one", 300, 200, 1),
    Pans(2, "tow", 200, 200, 2),
    Pans(3, "three", 100, 200, 3),
    Pans(4, "for", 200, 200, 1),
    Pans(5, "wer", 200, 200, 2),
    Pans(6, "fgt", 200, 200, 1),
]

pprint.pprint(group_by_keys(lp, ["width", "length"]))

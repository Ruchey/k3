# -*- coding: utf-8 -*-
# Вспомогательные функции

__author__ = 'Виноградов А.Г. г.Белгород июнь 2019'


def normkeyprop(lst):
    "В ключах убирает первые цифровые символы, пробел внутри слов удаляет"

    newlist = []
    for tup in lst:
        key = tup[0].replace(" ","").lower().lstrip('0123456789.- ')
        newlist.append((key, tup[1]))
    return newlist
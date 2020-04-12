"""Модуль для создания отчётов для программы К3-Мебель

xl -  модуль для работы с документами Excel
db -  модуль для подключения к базам данных access
base - модуль для получения информации некоторых таблиц базы
panel модуль для получения информации по панелям
nomenclature модуль работы с номенклатурой
profile модуль работы с профилями
long модуль для работы с длиномерами
"""

from . import base, db, xl, long, nomenclature, panel, prof, utils

__author__ = "Виноградов А.Г."
__copyright__ = "Copyright 2015, г.Белгород"
__license__ = "GPL"
__version__ = "1.0.0"
__maintainer__ = "Виноградов А.Г."
__email__ = "lvar-8@ya.ru"
__status__ = "Development"

__all__ = ['base', 'db', 'xl', 'long', 'nomenclature', 'panel', 'prof.py', 'utils']

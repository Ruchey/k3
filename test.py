# -*- coding: utf-8 -*-

from collections import OrderedDict, namedtuple
from  pprint import pprint

import k3r
from pprint import pprint

num = 1
fileDB = r'd:\К3\Самара\Самара черновик\{0}\{0}.mdb'.format(num)
projreppath = r'd:\К3\Самара\Самара черновик\{0}\Reports'.format(num)
project = "Деталировка"
db = k3r.db.DB()
db.open(fileDB)
pr = k3r.prof.Profile(db)
p = pr.profiles()
nm = k3r.nomenclature.Nomenclature(db)
bn = nm.bands()
ln = k3r.long.Long(db)

pprint(pr.profiles())
pprint(pr.total())
print()

db.close()
pass

# Вывод: unitpos, type, table, mat_id, goodsid
# 'type', 'matid', 'length', 'goodsid'

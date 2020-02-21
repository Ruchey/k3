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
acc = nm.acc_by_uid()
acl = nm.acc_long()
pprint(acc)
print()
pprint(acl)

db.close()
pass

# Вывод: unitpos, type, table, mat_id, goodsid
# 'type', 'matid', 'length', 'goodsid'

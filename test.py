# -*- coding: utf-8 -*-

import k3Report

from pprint import pprint


num = 5
fileDB = r'd:\К3\Самара\Самара черновик\{0}\{0}.mdb'.format(num)
projreppath = r'd:\К3\Самара\Самара черновик\{0}\Reports'.format(num)

db = k3Report.db.DB(fileDB)

pn = k3Report.panel.Panel(db)
nm = k3Report.nomenclature.Nomenclature(db)
lg = k3Report.longs.Longs(db)
bs = k3Report.base.Base(db)

res = bs.telems(2547)
#res = nm.properties(8426)

pprint(res)





db.close()


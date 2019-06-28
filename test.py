# -*- coding: utf-8 -*-

import k3Report

from pprint import pprint


num = 1
fileDB = r'd:\К3\Самара\Самара черновик\{0}\{0}.mdb'.format(num)
projreppath = r'd:\К3\Самара\Самара черновик\{0}\Reports'.format(num)

db = k3Report.db.DB(fileDB)

pn = k3Report.panel.Panel(db)
nm = k3Report.nomenclature.Nomenclature(db)
lg = k3Report.longs.Longs(db)
bs = k3Report.base.Base(db)

res = pn.slots_y_par(320)
a = "Паз по X {0}".format("; ".join(list(map('{0.beg}ш{0.width}г{0.depth}'.format, res))))
#res = nm.properties(8426)

pprint(a)




db.close()


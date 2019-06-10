# -*- coding: utf-8 -*-

import k3Report

from pprint import pprint



fileDB = r'd:\К3\Самара\Самара черновик\2\2.mdb'
projreppath = r'd:\К3\Самара\Самара черновик\2\Reports'

db = k3Report.db.DB(fileDB)
pn = k3Report.panel.Panel(db)

r1 = pn.cnt_drill_pans()

print(r1)




db.close()


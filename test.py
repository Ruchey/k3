# -*- coding: utf-8 -*-

import k3Report

from pprint import pprint

import time


fileDB = r'd:\К3\Самара\Самара черновик\1\1.mdb'
projreppath = r'd:\К3\Самара\Самара черновик\1\Reports'
startTime = time.time()
db = k3Report.db.DB(fileDB)
pn = k3Report.panel.Panel(db)

r1 = pn.cnt_holes_pan(608)
r2 = pn.cnt_holes_pan(738)
print(r1)
print(r2)


db.close()

print(time.time() - startTime)
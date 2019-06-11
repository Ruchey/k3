# -*- coding: utf-8 -*-

import k3Report

from pprint import pprint



fileDB = r'd:\К3\Самара\Самара черновик\1\1.mdb'
projreppath = r'd:\К3\Самара\Самара черновик\1\Reports'

db = k3Report.db.DB(fileDB)

pn = k3Report.panel.Panel(db)
nm = k3Report.nomenclature.Nomenclature(db)

matid = nm.properties(8528)

print(matid)




db.close()


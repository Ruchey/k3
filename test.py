# -*- coding: utf-8 -*-

import k3Report

from pprint import pprint


num = 46
fileDB = r'd:\К3\Самара\Самара черновик\{0}\{0}.mdb'.format(num)
projreppath = r'd:\К3\Самара\Самара черновик\{0}\Reports'.format(num)

db = k3Report.db.DB(fileDB)

pn = k3Report.panel.Panel(db)
nm = k3Report.nomenclature.Nomenclature(db)

matid = nm.bands(10)

pprint(matid)





db.close()


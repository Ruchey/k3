from pprint import pprint

import k3r
from k3r import get_tables

fl = r'd:\test.xlsx'
fileDB = r'd:\К3\Самара\Самара черновик\1\1.mdb'


db = k3r.db.DB(fileDB)

tb = get_tables.Specific(db)
pr = k3r.prof.Profile(db)
ln = k3r.long.Long(db)

# pprint(pr.total())
ln = ln.long_list()

pprint(ln)

# xl.save(fl)
# os.startfile(fl)


det = [645, 645, 645, 560, 560, 2100, 2100, 1800, 1800, 430, 125, 320, 320]
prep = 5300


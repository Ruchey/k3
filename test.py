import os
import k3r


fl = r'd:\test.xlsx'
fileDB = r'd:\К3\Самара\Самара черновик\2\2.mdb'


db = k3r.db.DB()
db.open(fileDB)

pn = k3r.panel.Panel(db)
bs = k3r.base.Base(db)

print(pn.par_bent_pan(454))





# xl.save(fl)
# os.startfile(fl)

import os, openpyxl as op, random as r

TUTOR_NAME_CELL = (3,3)
os.chdir(r"redacted")

names = open("redacted")
names = names.readlines()

fns = os.listdir()
names = r.sample(names, len(fns))
for fn in fns:
    wb = op.open(fn)
    ws = wb.active
    ws.cell(TUTOR_NAME_CELL[0],TUTOR_NAME_CELL[1]).value = names.pop()
    wb.save(fn)
    wb.close()
    
    
    




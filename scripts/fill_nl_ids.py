"""Fill column A (ID) of an NL shipping file from the Ship Cost Matrix.
Usage: fill_nl_ids.py MATRIX.xlsx NL.xlsx OUT.xlsx"""
import sys, openpyxl

def key(r):
    s = lambda v: (str(v).strip().upper() if v else '')
    return (s(r[1]), s(r[4]), s(r[5]), s(r[9]), s(r[10]), s(r[11]))

matrix, nl, out = sys.argv[1:4]
ids = {}
for r in openpyxl.load_workbook(matrix, read_only=True).active.iter_rows(min_row=2, values_only=True):
    if r[0] is not None:
        ids.setdefault(key(r), set()).add(r[0])

wb = openpyxl.load_workbook(nl); ws = wb.active
miss = amb = 0
for row in ws.iter_rows(min_row=2):
    vals = [c.value for c in row]
    if not vals[1] or not vals[10]:
        continue  # blank/template row
    found = ids.get(key(vals), set())
    if len(found) == 1:
        row[0].value = next(iter(found))
    else:
        miss += len(found) != 1 and not found; amb += len(found) > 1
        print('row', row[0].row, 'matches:', sorted(found), vals[1], vals[9], vals[10])
wb.save(out)
print('unmatched', miss, 'ambiguous', amb)

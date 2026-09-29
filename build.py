# -*- coding: utf-8 -*-
"""Packs prices.csv / china.csv into the Box Rate Desk page.
Run after extract.py.  Writes artifact-body.html (for publishing) and rate-desk.html (offline copy).
"""
import csv, json, os, datetime, collections

HERE = os.path.dirname(os.path.abspath(__file__))
NP, NO, CBC, OTH = -3, 0, -1, -2

def read(name):
    with open(os.path.join(HERE, name), encoding='utf-8-sig') as fh:
        return list(csv.DictReader(fh))

routes_raw = read('prices.csv')
china_raw  = read('china.csv')
try:
    notices_raw = read('notices.csv')
except FileNotFoundError:
    notices_raw = []

dates = sorted({r['date'] for r in routes_raw} | {r['date'] for r in china_raw})
di = {d: i for i, d in enumerate(dates)}
nd = len(dates)

loc, loc_i = [], {}
def L(en, zh):
    if en not in loc_i:
        loc_i[en] = len(loc); loc.append([en, zh or en])
    elif zh and loc[loc_i[en]][1] == en:
        loc[loc_i[en]][1] = zh
    return loc_i[en]

def idx(lst, v):
    if v not in lst: lst.append(v)
    return lst.index(v)

corr, box, grade, sub, stock = [], [], [], [], ['']
sub_of = {}

def val(r):
    p = r['price']
    if p:
        try: return float(p)
        except ValueError: pass
    s = (r['status'] or '').strip().lower()
    if s == 'not offered': return NO
    if s == 'case by case': return CBC
    return OTH if s else NP

series = collections.defaultdict(lambda: [NP] * nd)
stock_of = {}
for r in routes_raw:
    c = idx(corr, r['series']); o = L(r['origin'], r['origin_zh']); d = L(r['dest'], r['dest_zh'])
    b = idx(box, r['box']); g = idx(grade, r['grade'] or '')
    if r['sub']:
        sub_of[d] = idx(sub, r['sub'])
    k = (c, o, d, b, g)
    series[k][di[r['date']]] = val(r)
    if r['stock']:
        stock_of[k] = r['stock']

def encode(arr):
    out, prev = [], NP
    for i, v in enumerate(arr):
        if v != prev:
            out.append(i); out.append(round(v, 2) if v > 0 else v)
            prev = v
    return out

box_order = {'20GP': 0, '40GP': 1, '40HQ': 2, '45HQ': 3}
box.sort(key=lambda b: box_order.get(b, 9))
grade.sort(key=lambda g: (g != '', g))

out_routes = []
for (c, o, d, b, g), arr in series.items():
    pts = encode(arr)
    if not pts: continue
    st = stock_of.get((c, o, d, b, g), '')
    if st not in stock: stock.append(st)
    out_routes.append([c, o, d, box.index(box[b]) if False else b, g, sub_of.get(d), pts, stock.index(st)])
# box/grade were re-sorted after indexing -> remap
# (rebuild index maps from the original insertion order)
out_routes = []
box2, grade2 = [], []
for r in routes_raw[:0]: pass
box_list = sorted({r['box'] for r in routes_raw}, key=lambda b: box_order.get(b, 9))
grade_list = sorted({(r['grade'] or '') for r in routes_raw}, key=lambda g: (g != '', g))
for (c, o, d, b, g), arr in series.items():
    pts = encode(arr)
    if not pts: continue
    st = stock_of.get((c, o, d, b, g), '')
    if st not in stock: stock.append(st)
    out_routes.append([c, o, d, box_list.index(box[b]), grade_list.index(grade[g]),
                       sub_of.get(d), pts, stock.index(st)])

# --- china ---
cloc, cloc_i = [], {}
def CL(en, zh):
    if en not in cloc_i:
        cloc_i[en] = len(cloc); cloc.append([en, zh or en])
    return cloc_i[en]
cbox = sorted({r['box'] for r in china_raw}, key=lambda b: box_order.get(b, 9))
pur = collections.defaultdict(lambda: [NP] * nd)
ret = collections.defaultdict(lambda: [NP] * nd)
for r in china_raw:
    li = CL(r['location'], r['location_zh']); bi = cbox.index(r['box']); t = di[r['date']]
    def num(x):
        try: return float(x) if x else NP
        except ValueError: return NP
    pur[(li, bi)][t] = num(r['purchase'])
    ret[(li, bi)][t] = num(r['retail'])
crows = [[li, bi, encode(pur[(li, bi)]), encode(ret[(li, bi)])] for (li, bi) in sorted(pur)]
defaults = [cloc_i[n] for n in ('SHANGHAI', 'QINGDAO', 'GUANGZHOU', 'NINGBO') if n in cloc_i]

notices = []
for n in notices_raw:
    if n['date'] not in di:
        continue
    c = corr.index(n['series']) if n['series'] in corr else -1
    notices.append([di[n['date']], c, n['kind'], n['text']])
notices.sort()

emails = len({r['src'] for r in routes_raw} | {r['src'] for r in china_raw})
DATA = {
    'gen': datetime.date.today().strftime('%d %b %Y'),
    'emails': emails,
    'pubWeeks': nd,
    'dates': dates,
    'loc': loc, 'box': box_list, 'grade': grade_list, 'corr': corr, 'sub': sub,
    'subOf': {str(k): v for k, v in sub_of.items()},
    'stock': stock,
    'routes': out_routes,
    'china': {'loc': cloc, 'box': cbox, 'rows': crows, 'defaults': defaults},
    'notices': notices,
}
blob = json.dumps(DATA, ensure_ascii=False, separators=(',', ':'))

open(os.path.join(HERE, 'dataset.json'), 'w', encoding='utf-8').write(blob)

tpl = open(os.path.join(HERE, 'template.html'), encoding='utf-8').read()
body = tpl.replace('__DATA__', blob)
open(os.path.join(HERE, 'artifact-body.html'), 'w', encoding='utf-8').write(body)

title_line = '<title>Neptune Box Rate Desk</title>'
standalone = ('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
              '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
              + title_line + '\n</head>\n<body>\n'
              + body.replace(title_line + '\n', '') + '\n</body>\n</html>\n')
open(os.path.join(HERE, 'rate-desk.html'), 'w', encoding='utf-8').write(standalone)

print(f'notices {len(notices)}')
print(f'routes {len(out_routes)}  china rows {len(crows)}  weeks {nd}  emails {emails}')
print('data blob  %.2f MB' % (len(blob.encode()) / 1e6))
print('page       %.2f MB' % (len(standalone.encode()) / 1e6))

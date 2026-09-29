# -*- coding: utf-8 -*-
"""
Neptune container price extractor.
Parses every .eml quotation in the folder (HTML tables + XLSX attachments,
both the 2024-25 and 2026 layouts) into one tidy dataset.

Usage:  python3 extract.py [folder]
Writes: dataset.json / prices.csv / china.csv / extract_log.txt  (in ./ of this script)
"""
import email, glob, os, re, sys, json, io, csv, datetime, traceback
from email import policy
from bs4 import BeautifulSoup

HERE = os.path.dirname(os.path.abspath(__file__))
# Emails live beside the tool, either flat or (since the Sep 2026 migration) grouped in
# dated export subfolders. Search recursively so new export folders are picked up as-is.
_PARENT = os.path.dirname(HERE)
_MAILDIR = os.path.join(_PARENT, 'Container Price Emails')
FOLDER = sys.argv[1] if len(sys.argv) > 1 else (_MAILDIR if os.path.isdir(_MAILDIR) else _PARENT)

BOX_RE = re.compile(r'^(20|40|45)\s*(GP|HQ|HC|OT|FR|RF|NOR)$', re.I)
DATE_IN_NAME = re.compile(r'(20\d{6})')
NUM_RE = re.compile(r'^\s*(?:US\s*\$|USD|\$)?\s*([\d,]+(?:\.\d+)?)\s*$')

SERIES = [
    ('Russia',       ['russia', '中俄', '俄罗斯']),
    ('Europe',       ['europe', '中欧', '欧洲']),
    ('Central Asia', ['asia', 'centra', '中亚']),
    ('China',        ['china', '国内']),
]

def series_of(subject, fname):
    s = (subject or '') + ' ' + fname
    sl = s.lower()
    for name, keys in SERIES:
        for k in keys:
            if k in sl:
                return name
    return None

TABLE_SERIES = [
    ('China',        ['国内', 'prices for china']),
    ('Russia',       ['中俄', 'for russia']),
    ('Europe',       ['中欧', '欧洲', 'for europe']),
    ('Central Asia', ['哈萨克斯坦', 'kazakh', '乌兹别克斯坦', 'uzbek', '独联体',
                      'other cis', '中亚', 'central asia', 'for asia']),
]
SUB_LABEL = [
    ('Kazakhstan', ['哈萨克斯坦', 'kazakh']),
    ('Uzbekistan', ['乌兹别克斯坦', 'uzbek']),
    ('Other CIS',  ['独联体', 'other cis']),
]

def _match(text, table):
    tl = text.lower()
    for name, keys in table:
        for k in keys:
            if k in tl:
                return name
    return None

def series_from_table(grid, sheet, default):
    head = ' '.join(' '.join(r) for r in grid[:3])[:600] + ' ' + (sheet or '')
    return _match(head, TABLE_SERIES) or default

def sub_from_table(grid, sheet):
    head = ' '.join(' '.join(r) for r in grid[:3])[:600] + ' ' + (sheet or '')
    return _match(head, SUB_LABEL) or ''

def norm(t):
    return ' '.join((t or '').replace('\xa0', ' ').split())

ALIAS = {
    'KARAGANDA': 'KARAGANDY',
    'MALASZEWICAE': 'MALASZEWICZE',
    'KHAVAROVSK': 'KHABAROVSK',
    'HABAROVSK': 'KHABAROVSK',
    'RYBACHYE': 'BALYKCHY',          # same city, renamed in later sheets
}
GRADE_RE = re.compile(r'[（(]\s*(适货标准|适货|CW|IICL)\s*[)）]?', re.I)
GRADE_MAP = {'适货': 'CW', '适货标准': 'CW', 'CW': 'CW', 'IICL': 'IICL'}
BAD_HDR = re.compile(r'箱况|Condition|备注|P\.S|说明|标准|Remark', re.I)

def split_name(cell):
    """'塔什干(CW) TASHKENT' -> ('塔什干','TASHKENT','CW')"""
    c = norm(cell)
    grade = ''
    m = GRADE_RE.search(c)
    if m:
        grade = GRADE_MAP.get(m.group(1).upper() if m.group(1).isascii() else m.group(1), '')
        c = GRADE_RE.sub(' ', c)
    c = norm(c.replace('（', ' ').replace('）', ' '))
    c = re.sub(r'\(\s*\)', ' ', c)
    m = re.match(r"^([^A-Za-z]*)\s*([A-Za-z][A-Za-z0-9 ,./&'\-\(\)\.]*)$", c)
    if m and m.group(2).strip():
        zh = m.group(1).strip(' ()（）')
        en = m.group(2).strip(' ()')
    else:
        zh = en = c.strip(' ()')
    en = ALIAS.get(en.upper(), en.upper()) if en else en
    return zh, en, grade

def parse_price(cell):
    c = norm(cell)
    if not c or c in {'/', '-', '--', '_'}:
        return None, 'not offered'
    m = NUM_RE.match(c)
    if m:
        try:
            v = float(m.group(1).replace(',', ''))
        except ValueError:
            return None, c
        if v <= 0:
            return None, c
        return v, None
    return None, c            # 'case by case', 'N/A', 'no service', etc.

# ---------- grid builders ----------
def grid_from_html_table(tbl):
    grid, r = {}, 0
    for tr in tbl.find_all('tr'):
        if tr.find_parent('table') is not tbl:
            continue
        c = 0
        for td in tr.find_all(['td', 'th']):
            if td.find_parent('table') is not tbl:
                continue
            while (r, c) in grid:
                c += 1
            rs = int(td.get('rowspan', 1) or 1)
            cs = int(td.get('colspan', 1) or 1)
            txt = norm(td.get_text(' ', strip=True))
            for i in range(min(rs, 60)):
                for j in range(min(cs, 60)):
                    grid[(r + i, c + j)] = txt
            c += cs
        r += 1
    if not grid:
        return []
    R = max(k[0] for k in grid) + 1
    C = max(k[1] for k in grid) + 1
    return [[grid.get((i, j), '') for j in range(C)] for i in range(R)]

def grid_from_sheet(ws):
    maxr = min(ws.max_row, 400)
    maxc = min(ws.max_column, 80)
    g = [['' for _ in range(maxc)] for _ in range(maxr)]
    for row in ws.iter_rows(min_row=1, max_row=maxr, max_col=maxc):
        for cell in row:
            v = cell.value
            if v is None:
                continue
            if isinstance(v, datetime.datetime):
                v = v.strftime('%Y-%m-%d')
            g[cell.row - 1][cell.column - 1] = norm(str(v))
    for rng in ws.merged_cells.ranges:
        r1, c1, r2, c2 = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        if r1 > maxr or c1 > maxc:
            continue
        v = g[r1 - 1][c1 - 1]
        for i in range(r1 - 1, min(r2, maxr)):
            for j in range(c1 - 1, min(c2, maxc)):
                g[i][j] = v
    return g

# ---------- table interpretation ----------
def find_update_date(grid):
    for row in grid[:6]:
        for j, cell in enumerate(row):
            if '更新日期' in cell or 'Update time' in cell:
                for k in range(j, min(j + 5, len(row))):
                    v = norm(row[k])
                    if v and '更新' not in v and 'Update' not in v:
                        for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d-%b-%Y', '%Y/%m/%d', '%Y%m%d'):
                            try:
                                return datetime.datetime.strptime(v.split(' 00:')[0].strip(), fmt).date()
                            except ValueError:
                                pass
                        m = DATE_IN_NAME.search(v.replace('-', '').replace('/', ''))
                        if m:
                            try:
                                return datetime.datetime.strptime(m.group(1), '%Y%m%d').date()
                            except ValueError:
                                pass
    return None

def parse_route_grid(grid):
    """Locates the column roles rather than assuming them: Jeff's tables vary
    (some carry an Available-Stock column, some a Container-Conditions column,
    destinations start at whichever column follows).
    Returns (dests, rows) with rows = (origin, box, stock, grade_override, row)."""
    hdr = None
    for i, row in enumerate(grid[:12]):
        j = ' '.join(row[:5])
        if '起租地' in j or 'On-hire' in j or '提箱地' in j:
            hdr = i
            break
    if hdr is None:
        return None, []
    ncol = max(len(r) for r in grid[hdr:hdr + 3]) if len(grid) > hdr else 0

    def htxt(c, upto):
        return ' '.join(grid[i][c] for i in range(hdr, upto) if c < len(grid[i]))

    roles = {}
    for c in range(min(ncol, 8)):
        t = htxt(c, min(hdr + 3, len(grid)))
        if 'origin' not in roles and ('起租地' in t or '提箱地' in t or 'On-hire' in t):
            roles['origin'] = c; continue
        if 'box' not in roles and ('箱型' in t or 'Type of Equipment' in t):
            roles['box'] = c; continue
        if 'stock' not in roles and ('可供库存' in t or 'Available Stock' in t):
            roles['stock'] = c; continue
    for c in range(ncol):
        if '箱况' in htxt(c, min(hdr + 3, len(grid))) or 'Container Condition' in htxt(c, min(hdr + 3, len(grid))):
            roles['cond'] = c
    oc, bc = roles.get('origin', 0), roles.get('box', 1)

    data_start = None
    for i in range(hdr + 1, len(grid)):
        v = norm(grid[i][bc]) if bc < len(grid[i]) else ''
        if BOX_RE.match(v.replace(' ', '')):
            data_start = i
            break
    if data_start is None:
        return None, []

    dests = []
    first = max(bc, roles.get('stock', -1)) + 1
    for c in range(first, ncol):
        if c == roles.get('cond'):
            continue
        name = ''
        for i in range(hdr, data_start):
            t = norm(grid[i][c]) if c < len(grid[i]) else ''
            if not t or '退租地' in t or '还箱地' in t or 'Off-hire' in t:
                continue
            name = t
        if not name or len(name) > 45 or BAD_HDR.search(name) or not re.search(r'[A-Za-z]{3}', name):
            continue
        dests.append((c, name))
    if not dests:
        return None, []

    rows, origin = [], None
    for row in grid[data_start:]:
        c0 = norm(row[oc]) if oc < len(row) else ''
        c1 = norm(row[bc]) if bc < len(row) else ''
        if c0.startswith('P.S') or c0.startswith('备注'):
            break
        if c0:
            origin = c0
        if not BOX_RE.match(c1.replace(' ', '')) or not origin:
            continue
        stock = norm(row[roles['stock']]) if 'stock' in roles and roles['stock'] < len(row) else ''
        cond = norm(row[roles['cond']]) if 'cond' in roles and roles['cond'] < len(row) else ''
        g = 'CW' if re.search(r'适货|Cargo worthy', cond, re.I) else ('IICL' if 'IICL' in cond.upper() else '')
        rows.append((origin, c1.upper().replace(' ', ''), stock, g, row))
    return dests, rows

def parse_china_grid(grid):
    hdr = None
    for i, row in enumerate(grid[:10]):
        if '提箱地' in ' '.join(row[:3]):
            hdr = i
            break
    if hdr is None:
        return []
    start = hdr + 1
    for i in range(hdr, min(hdr + 4, len(grid))):
        if 'Pick-up' in ' '.join(grid[i][:3]) or 'Type of' in ' '.join(grid[i][:3]):
            start = i + 1
    out, loc = [], None
    for row in grid[start:]:
        c0 = norm(row[0])
        c1 = norm(row[1]) if len(row) > 1 else ''
        if c0.startswith('P.S') or c0.startswith('备注'):
            break
        if c0:
            loc = c0
        if not BOX_RE.match(c1.replace(' ', '')) or not loc:
            continue
        g = lambda i: norm(row[i]) if len(row) > i else ''
        out.append((loc, c1.upper().replace(' ', ''), g(2), g(3), g(4), g(5)))
    return out

# ---------- per-file parsing (cached) ----------
CACHE = os.path.join(HERE, '.cache')
CACHE_V = 4          # bump to invalidate parsed-email cache

VOID_RE = re.compile(r'失效|作废|暂停|停止|no longer valid|immediately invalid|suspend', re.I)
REVISED_RE = re.compile(r'更正|更新|紧急|立即生效|即时调整|已调整|effective immediately|updat|correct', re.I)
CUT = ('更新日期', '有效期至', '起租地', '提箱地', '关于客户买箱的注意事项', '近期',
       'Update time', 'Valid Until', '史集青', '保密声明', 'CONFIDENTIALITY',
       'NEPTUNE LOGISTICS', '地址/Add', '箱管部')

def lead_text(msg):
    body = msg.get_body(preferencelist=('plain',))
    t = body.get_content() if body is not None else ''
    t = t.split('Best regards')[0]
    for c in CUT:
        i = t.find(c)
        if i > 0:
            t = t[:i]
    t = re.sub(r'^\s*Dear\s+all[，,]?\s*', '', t, flags=re.I)
    return ' '.join(t.split())[:400]

def classify(t):
    if VOID_RE.search(t): return 'void'
    if REVISED_RE.search(t): return 'revised'
    return 'note'
RCOLS = ['date','series','sub','origin_zh','origin','dest_zh','dest','grade','box','price','status','stock']
CCOLS = ['date','location_zh','location','box','purchase_last','purchase','retail','status','stock']

def parse_file(path):
    base = os.path.basename(path)
    out = {'routes': [], 'china': [], 'log': []}
    try:
        with open(path, encoding='utf-8', errors='replace') as fh:
            msg = email.message_from_file(fh, policy=policy.default)
    except Exception as e:
        out['log'].append(f'ERROR open {base}: {e}')
        return out
    subject = msg.get('subject', '')
    ser = series_of(subject, base)
    if not ser:
        out['log'].append(f'SKIP no series: {base}')
        return out
    m = DATE_IN_NAME.search(base)
    fallback_date = datetime.datetime.strptime(m.group(1), '%Y%m%d').date() if m else None

    if re.match(r'^\d+_Re', base) and fallback_date:
        txt = lead_text(msg)
        if txt:
            out['notice'] = [fallback_date.isoformat(), ser, classify(txt), txt]

    grids = []
    try:
        body = msg.get_body(preferencelist=('html',))
        if body is not None:
            soup = BeautifulSoup(body.get_content(), 'lxml')
            for t in soup.find_all('table'):
                if t.find('table'):
                    continue
                g = grid_from_html_table(t)
                if len(g) < 4:
                    continue
                flat = ' '.join(' '.join(r[:4]) for r in g[:8])
                if '起租地' in flat or 'On-hire' in flat or '提箱地' in flat:
                    grids.append((g, ''))
    except Exception as e:
        out['log'].append(f'ERROR html {base}: {e}')

    for part in msg.walk():
        fn = part.get_filename()
        if not fn or not fn.lower().endswith(('.xlsx', '.xlsm')):
            continue
        try:
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(part.get_payload(decode=True)), data_only=True)
            for ws in wb.worksheets:
                g = grid_from_sheet(ws)
                if len(g) < 4:
                    continue
                flat = ' '.join(' '.join(r[:4]) for r in g[:8])
                if '起租地' in flat or 'On-hire' in flat or '提箱地' in flat:
                    grids.append((g, norm(ws.title)))
        except Exception as e:
            out['log'].append(f'ERROR xlsx {base} [{fn}]: {e}')

    if not grids:
        out['log'].append(f'NO TABLE: {base}')
        return out

    for g, sheet in grids:
        qdate = find_update_date(g) or fallback_date
        if not qdate:
            out['log'].append(f'NO DATE: {base}')
            continue
        ds = qdate.isoformat()
        tser = series_from_table(g, sheet, ser)
        if tser != ser:
            out['log'].append(f'SERIES OVERRIDE {base} [{sheet}] {ser} -> {tser}')
        sub = sub_from_table(g, sheet)
        if tser == 'China':
            for loc, box, last_w, this_w, retail, stock in parse_china_grid(g):
                zh, en, _ = split_name(loc)
                pw, sw = parse_price(this_w)
                pl, _sl = parse_price(last_w)
                pr, sr = parse_price(retail)
                out['china'].append([ds, zh, en, box, pl, pw, pr, sw or sr or '', stock])
        else:
            dests, rows = parse_route_grid(g)
            if not dests:
                out['log'].append(f'NO HEADER: {base} [{sheet}]')
                continue
            for origin, box, stock, grade_col, row in rows:
                ozh, oen, _ = split_name(origin)
                for j, dname in dests:
                    if j >= len(row):
                        continue
                    dzh, den, grade = split_name(dname)
                    grade = grade or grade_col
                    price, status = parse_price(row[j])
                    out['routes'].append([ds, tser, sub, ozh, oen, dzh, den, grade, box, price, status or '', stock])
    return out

def cached(path):
    base = os.path.basename(path)
    rel = os.path.relpath(path, FOLDER)
    st = os.stat(path)
    cp = os.path.join(CACHE, base + '.json')
    if os.path.exists(cp):
        try:
            d = json.load(open(cp, encoding='utf-8'))
            if (d.get('m') == int(st.st_mtime) and d.get('s') == st.st_size
                    and d.get('v') == CACHE_V and d.get('p', rel) == rel):
                return d, True
        except Exception:
            pass
    d = parse_file(path)
    d['m'] = int(st.st_mtime); d['s'] = st.st_size; d['v'] = CACHE_V; d['p'] = rel
    os.makedirs(CACHE, exist_ok=True)
    with open(cp, 'w', encoding='utf-8') as fh:
        json.dump(d, fh, ensure_ascii=False, separators=(',', ':'))
    return d, False

# ---------- main ----------
BUDGET = float(os.environ.get('EXTRACT_BUDGET', '95'))

def main():
    import time
    t0 = time.time()
    files = sorted(glob.glob(os.path.join(FOLDER, '**', '*.eml'), recursive=True))
    parsed, fresh, remaining = [], 0, 0
    for k, path in enumerate(files):
        if time.time() - t0 > BUDGET:
            remaining = len(files) - k
            break
        d, was_cached = cached(path)
        if not was_cached:
            fresh += 1
        parsed.append((path, d))
    if remaining:
        print(f'PARTIAL: parsed {len(parsed)}/{len(files)} ({fresh} newly) — {remaining} left, run again')
        return 1

    log, route_recs, china_recs, with_tables, notices = [], {}, {}, 0, []
    for path, d in parsed:
        base = os.path.basename(path)
        # Jeff's later message is authoritative: "更正报价，以此为准" replies correct the
        # original send, so the highest email id wins regardless of Re: status.
        prio = int(base.split('_')[0]) if base.split('_')[0].isdigit() else 0
        log.extend(d.get('log', []))
        if d.get('notice'):
            notices.append(d['notice'] + [base])
        if d['routes'] or d['china']:
            with_tables += 1
        for r in d['routes']:
            rec = dict(zip(RCOLS, r)); rec['src'] = base
            key = (rec['date'], rec['series'], rec['origin'], rec['dest'], rec['box'], rec['grade'])
            prev = route_recs.get(key)
            if prev and prev['_p'] > prio:
                continue
            rec['_p'] = prio; route_recs[key] = rec
        for r in d['china']:
            rec = dict(zip(CCOLS, r)); rec['src'] = base
            key = (rec['date'], rec['location'], rec['box'])
            prev = china_recs.get(key)
            if prev and prev['_p'] > prio:
                continue
            rec['_p'] = prio; china_recs[key] = rec

    routes = sorted(route_recs.values(), key=lambda r: (r['date'], r['series'], r['origin'], r['dest'], r['box'], r['grade']))
    chinas = sorted(china_recs.values(), key=lambda r: (r['date'], r['location'], r['box']))
    for r in routes: r.pop('_p', None)
    for r in chinas: r.pop('_p', None)

    with open(os.path.join(HERE, 'prices.csv'), 'w', newline='', encoding='utf-8-sig') as fh:
        w = csv.DictWriter(fh, fieldnames=RCOLS + ['src']); w.writeheader(); w.writerows(routes)
    with open(os.path.join(HERE, 'china.csv'), 'w', newline='', encoding='utf-8-sig') as fh:
        w = csv.DictWriter(fh, fieldnames=CCOLS + ['src']); w.writeheader(); w.writerows(chinas)

    notices.sort()
    with open(os.path.join(HERE, 'notices.csv'), 'w', newline='', encoding='utf-8-sig') as fh:
        w = csv.writer(fh); w.writerow(['date', 'series', 'kind', 'text', 'src']); w.writerows(notices)

    dates = sorted({r['date'] for r in routes} | {r['date'] for r in chinas})
    summary = [
        f'source folder      : {FOLDER}',
        f'files scanned      : {len(files)}  ({fresh} newly parsed, rest from cache)',
        f'files with tables  : {with_tables}',
        f'route rows         : {len(routes)}',
        f'china rows         : {len(chinas)}',
        f'distinct dates     : {len(dates)}  ({dates[0] if dates else "-"} .. {dates[-1] if dates else "-"})',
        f'mid-week notices   : {len(notices)}  ({sum(1 for n in notices if n[2] == "void")} withdrawals)',
    ]
    with open(os.path.join(HERE, 'extract_log.txt'), 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(summary) + '\n\n' + '\n'.join(log) + '\n')
    print('\n'.join(summary))
    print(f'issues logged: {len(log)}')
    return 0

if __name__ == '__main__':
    raise SystemExit(main())

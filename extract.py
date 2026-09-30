# -*- coding: utf-8 -*-
"""
Rail price extractor  (NEP Rail & Container Price Portal)
--------------------------------------------------------
Scans every export folder under "Rail Freight", pulls China -> Europe rail
prices for the ORIGINS and DESTINATIONS listed below, keeps SOC and COC
separate, picks the latest version of each quote date, and writes:

    rail-price-tool/rail_prices.csv          one row per price (clean table)
    rail-price-tool/rail_prices.json         same data, used by the dashboard

To add a destination (e.g. Hamburg), add a line to DESTINATIONS and re-run.
    rail-price-tool/extract_log.txt          files used / skipped / warnings

Usage:   python extract.py                 (needs openpyxl + xlrd)
         python extract.py --budget 150    (stop after N seconds; re-run to continue)
Parsed workbooks are cached in rail-price-tool/.cache so re-runs are fast.
"""
import os, re, sys, json, gzip, hashlib, time, csv, datetime, warnings
warnings.filterwarnings('ignore')

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                      # the "Rail Freight" folder
CACHE = os.path.join(HERE, '.cache')
BUDGET = float(sys.argv[sys.argv.index('--budget') + 1]) if '--budget' in sys.argv else 1e9
YEAR = 2026

FAMILIES = [  # (key, filename test)
    ('ZY',  lambda b: '欧洲自营价格表' in b),
    ('FZY', lambda b: '欧洲非自营价格表' in b),
    ('CY',  lambda b: '成都（城厢站）欧线' in b),
    ('KLH', lambda b: '跨里海' in b),
]
FAMILY_LABEL = {'ZY': '欧洲自营价格表', 'FZY': '欧洲非自营价格表', 'CY': '成都城厢欧线/重庆非自营', 'KLH': 'NEP跨里海报价表'}
ORIGINS = [  # (regex on origin/route text, canonical name)
    (r'成都|城厢|Chengdu', 'Chengdu (Chengxiang)'),
    (r'重庆|团结村|Chongqing', 'Chongqing (Tuanjiecun)'),
    (r"西安|Xi'?an", "Xi'an (Intl Port)"),
    (r'武汉|Wuhan', 'Wuhan'),
    (r'长沙|Changsha', 'Changsha'),
    (r'南昌|赣州|宜春|萍乡|新余|吉安|九江|株洲', 'Jiangxi/Hunan (near Yichun)'),
]
DESTINATIONS = [  # name shown in the dashboard, Chinese name used in the sheets, regex matched on the destination cell
    {'name': 'Budapest', 'cn': '布达佩斯', 'pattern': r'布达佩斯|budapest'},
    # {'name': 'Hamburg',  'cn': '汉堡',     'pattern': r'汉堡|hamburg'},
    # {'name': 'Duisburg', 'cn': '杜伊斯堡', 'pattern': r'杜伊斯堡|杜堡|duisburg'},
]
DEST = re.compile('|'.join('(?:%s)' % d['pattern'] for d in DESTINATIONS), re.I)


def dest_of(text):
    for d in DESTINATIONS:
        if re.search(d['pattern'], text or '', re.I):
            return d
    return None


# ---------------------------------------------------------------- reading
def _conv(v):
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        return v.isoformat()
    return v


def load(path):
    rel = os.path.relpath(path, ROOT)
    key = hashlib.md5((rel + str(os.path.getmtime(path))).encode()).hexdigest()
    cp = os.path.join(CACHE, key + '.json.gz')
    if os.path.exists(cp):
        return json.load(gzip.open(cp, 'rt', encoding='utf-8'))
    sheets = []
    if path.lower().endswith('.xlsx'):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        for ws in wb.worksheets:
            sheets.append({'name': ws.title, 'state': ws.sheet_state,
                           'rows': [[_conv(c) for c in r] for r in ws.iter_rows(values_only=True)]})
    else:
        import xlrd
        wb = xlrd.open_workbook(path)
        for sh in wb.sheets():
            rows = []
            for i in range(sh.nrows):
                r = []
                for c in sh.row(i):
                    v = c.value
                    if c.ctype == 3:
                        try:
                            v = xlrd.xldate.xldate_as_datetime(v, wb.datemode).isoformat()
                        except Exception:
                            pass
                    elif c.ctype in (0, 6):
                        v = None
                    r.append(v)
                rows.append(r)
            sheets.append({'name': sh.name, 'state': ['visible', 'hidden', 'veryHidden'][sh.visibility], 'rows': rows})
    d = {'path': rel, 'mtime': os.path.getmtime(path), 'sheets': sheets}
    os.makedirs(CACHE, exist_ok=True)
    json.dump(d, gzip.open(cp, 'wt', encoding='utf-8'), ensure_ascii=False, default=str)
    return d


# ---------------------------------------------------------------- helpers
def s(v):
    return '' if v is None else str(v).strip()


def flat(v):
    return re.sub(r'\s+', ' ', s(v))


def canon_origin(text):
    for pat, name in ORIGINS:
        if re.search(pat, text or '', re.I):
            return name
    return None


def parse_price(v):
    """cell -> (number or None, leftover note)"""
    if v is None:
        return None, ''
    if isinstance(v, (int, float)):
        return (float(v), '') if v >= 500 else (None, '')
    t = flat(v)
    if not t or t in ('/', '-', '无', '#VALUE!', '#REF!'):
        return None, ''
    m = re.search(r'(\d{4,6}(?:\.\d+)?)', t.replace(',', ''))
    if not m:
        return None, t
    note = (t[:m.start()] + t[m.end():]).strip(' /')
    return float(m.group(1)), note


def file_date(b):
    m = re.search(r'NEP-RPP-(\d{4})-(\d{2})(\d{2})', b)
    if m:
        return f'{m.group(1)}-{m.group(2)}-{m.group(3)}'
    m = re.search(r'(?<![\d.])(\d{1,2})\.(\d{1,2})(?=\s*(?:\(\d+\))?\s*\.xls)', b)
    if m:
        return f'{YEAR}-{int(m.group(1)):02d}-{int(m.group(2)):02d}'
    m = re.search(r'(20\d{2})(\d{2})(\d{2})', b)
    if m:
        return f'{m.group(1)}-{m.group(2)}-{m.group(3)}'
    return None


def sheet_date(name):
    m = re.search(r'(?<!\d)(\d{2})\.(\d{1,2})(?!\d)', name)
    if m and 1 <= int(m.group(1)) <= 12:
        return f'{YEAR}-{int(m.group(1)):02d}-{int(m.group(2)):02d}'
    return None


def dep_month(b, sheetname, title):
    for t in (sheetname, b, title):
        m = re.search(r'(\d{1,2}(?:-\d{1,2})?)月', t or '')
        if m:
            return m.group(1) + '月'
    return ''


def service_tags(*texts):
    t = ' '.join(texts)
    tags = []
    for k, lab in [('快线', 'Express 快线'), ('时刻表', 'Timetable 时刻表'), ('全时', 'Timetable 时刻表'),
                   ('普线', 'Regular 普线'), ('图定', 'Scheduled 图定'), ('不限制', 'Unrestricted 不限制'),
                   ('定制', 'Custom 定制')]:
        if k in t and lab not in tags:
            tags.append(lab)
    return tags


ADDER = re.compile(r'([^，,；;。|]{0,18}?)(加收|\+|另收|需加)\s*(RMB|USD|CNY|美金)?\s*(\d{2,5})\s*(RMB|USD|CNY|元|美金)?\s*(/柜|/票|/40HQ|/Bill)?')
EXCL = re.compile(r'不含([^，,；;。|（(]{0,10}?)\s*(\d{2,5})\s*(RMB|USD|元)\s*(/柜|/票)?')


def adders(text):
    """Surcharges written in the remark text. 'excluded': the sheet says the fee is NOT included (不含)."""
    t = flat(text)
    out = []
    for m in EXCL.finditer(t):
        out.append({'ctx': ('不含' + m.group(1)).strip(), 'amount': float(m.group(2)),
                    'cur': m.group(3).replace('元', 'RMB'), 'per': (m.group(4) or '/柜').strip('/'), 'excluded': True})
    for m in ADDER.finditer(t):
        cur = (m.group(3) or m.group(5) or '').replace('美金', 'USD').replace('元', 'RMB').replace('CNY', 'RMB')
        if not cur:
            continue
        amt = float(m.group(4))
        ctx = (m.group(1) or '').strip()
        if amt < 30 or ctx.endswith('提'):          # pickup-location lists are handled by coc_pickup()
            continue
        out.append({'ctx': ctx[-18:], 'amount': amt, 'cur': cur,
                    'per': (m.group(6) or '').strip('/'), 'excluded': False})
    seen, res = set(), []
    for a in out:
        k = (a['amount'], a['cur'])
        if k not in seen:
            seen.add(k)
            res.append(a)
    return res[:10]


ORIGIN_CITY = {'Chengdu (Chengxiang)': '成都', 'Chongqing (Tuanjiecun)': '重庆', "Xi'an (Intl Port)": '西安',
               'Wuhan': '武汉', 'Changsha': '长沙'}


def coc_pickup(remark, cn='布达佩斯'):
    """Parse the Budapest clause of the COC box-pickup rules, e.g.
    '布达佩斯：上海、宁波提箱价；广州黄埔提+337RMB，武汉提+675RMB；西安、成都提+2025RMB；'
    -> {'base': [...], 'adders': {'广州黄埔': 337, ...}, 'text': clause}"""
    t = flat(remark)
    m = re.search(re.escape(cn) + r'[^：:，,；;]{0,4}[：:]', t)
    if not m:
        return None
    rest = t[m.end():]
    nxt = re.search(r'[；;。|]\s*[\u4e00-\u9fa5A-Za-z（）()]{2,10}[：:]', rest)
    clause = rest[:nxt.start()] if nxt else rest[:300]
    base, add = [], {}
    for part in re.split(r'[；;，,]', clause):
        part = part.strip()
        if not part:
            continue
        mm = re.match(r'(.+?)提\s*[+＋]\s*(\d{2,5})\s*(RMB|元|USD)?', part)
        if mm:
            for c in re.split(r'[、/]', mm.group(1)):
                if c.strip():
                    add[c.strip()] = (float(mm.group(2)), 'USD' if mm.group(3) == 'USD' else 'RMB')
            continue
        mm = re.match(r'(.+?)提箱价', part)
        if mm:
            base += [c.strip() for c in re.split(r'[、/]', mm.group(1)) if c.strip()]
    return {'base': base, 'adders': add, 'text': (cn + '：' + clause)[:300]}


# ---------------------------------------------------------------- table parser (ZY / FZY / CY)
def classify_header(h):
    h = flat(h)
    if not re.search(r'SOC|COC', h, re.I) or re.search(r'询价|联络', h):
        return None
    if re.match(r'^\d{1,2}\.\d{1,2}', h):          # one-off dated column, e.g. "8.30站到站SOC"
        return None
    box = 'COC' if 'COC' in h.upper() else 'SOC'
    if re.search(r'20\s*HQ\s*一组|2\*20', h):
        size = '2x20ft'
    elif re.search(r'20\s*(HQ|GP|呎|尺)', h):
        size = '20ft'
    elif re.search(r'40\s*GP', h):
        size = '40GP'
    elif re.search(r'40\s*HQ', h):
        size = '40HQ'
    else:
        size = None
    cur = 'USD' if re.search(r'USD|美金', h, re.I) else ('RMB' if re.search(r'RMB|人民币', h) else None)
    return {'box': box, 'size': size, 'cur': cur, 'ref': bool(re.search(r'参考', h)), 'hdr': h}


def find_headers(rows):
    hs = []
    for i, r in enumerate(rows):
        cells = [flat(v) for v in r]
        dcol = next((j for j, c in enumerate(cells) if re.match(r'^(目的站|线路)', c)), None)
        if dcol is None or not any(re.search(r'SOC|COC', c) for c in cells):
            continue
        prices = {j: k for j, c in enumerate(cells) for k in [classify_header(c)] if k}
        sizes = [k['size'] for k in prices.values() if k['size'] and k['size'] != '2x20ft']
        dflt = '40HQ' if ('40HQ' in sizes or not sizes) else max(set(sizes), key=sizes.count)
        for k in prices.values():
            if not k['size']:
                k['size'] = dflt
        hs.append({'row': i, 'dcol': dcol,
                   'ocol': next((j for j, c in enumerate(cells) if re.match(r'^(发运站|发站)', c)), None),
                   'scol': next((j for j, c in enumerate(cells) if re.match(r'^(班期|ETD)', c)), None),
                   'pcol': next((j for j, c in enumerate(cells) if re.match(r'^口岸', c)), None),
                   'rcols': [j for j, c in enumerate(cells) if c.startswith('备注')],
                   'prices': prices})
    return hs


def parse_table_sheet(fam, sh):
    rows = sh['rows']
    out = []
    title = ' '.join(flat(v) for r in rows[:3] for v in r if v)[:400]
    hdrs = find_headers(rows)
    for hi, H in enumerate(hdrs):
        end = hdrs[hi + 1]['row'] if hi + 1 < len(hdrs) else len(rows)
        sect = ' '.join(flat(v) for v in rows[H['row'] - 1] if v) if H['row'] > 0 else ''
        origin_txt, sched, port, remarks = '', '', '', {}
        for i in range(H['row'] + 1, end):
            r = list(rows[i]) + [None] * 30
            if H['ocol'] is not None and s(r[H['ocol']]):
                origin_txt, sched, port, remarks = flat(r[H['ocol']]), '', '', {}
            if H['scol'] is not None and s(r[H['scol']]):
                sched = flat(r[H['scol']])
            if H['pcol'] is not None and s(r[H['pcol']]):
                port = flat(r[H['pcol']])
            for rc in H['rcols']:
                if s(r[rc]):
                    remarks[rc] = flat(r[rc])
            dest_txt = flat(r[H['dcol']])
            dst = dest_of(dest_txt)
            if not dst:
                continue
            org = canon_origin(origin_txt) or canon_origin(dest_txt)
            dt = dest_txt.strip()
            if fam == 'CY' and dt.startswith('重庆'):
                org = 'Chongqing (Tuanjiecun)'
            if fam == 'CY' and dt.startswith('城厢'):
                org = 'Chengdu (Chengxiang)'
            if not org:
                continue
            if fam == 'CY':
                operator = '非自营 3rd-party' if ('非自营' in origin_txt or org.startswith('Chongqing')) else '自营 own'
            else:
                operator = '非自营 3rd-party' if fam == 'FZY' else '自营 own'
            product = 'South China door-to-station 华南门到站' if '门到站' in sh['name'] else 'Station-to-station 站到站'
            remark = ' | '.join(remarks[k] for k in sorted(remarks))[:900]
            tags = service_tags(dest_txt, port)
            for j, k in H['prices'].items():
                val, note = parse_price(r[j])
                if val is None:
                    continue
                cur = k['cur'] or ('USD' if (re.search(r'美金|USD', sect) or fam == 'FZY') else 'RMB')
                if cur == 'RMB' and val < 20000:        # USD-sized number under an RMB header
                    cur = 'USD'
                out.append({'destination': dst['name'], 'box': k['box'], 'size': k['size'], 'currency': cur, 'amount': val,
                            'is_reference': k['ref'], 'price_note': note[:80], 'origin': org,
                            'origin_raw': origin_txt[:60], 'dest_raw': dest_txt[:80], 'port': port[:40],
                            'schedule': sched[:160], 'service': tags, 'operator': operator, 'product': product,
                            'remark': remark, 'header': k['hdr'][:60], 'row': i + 1, 'col': j + 1, 'title': title})
    return out


# ---------------------------------------------------------------- trans-Caspian parser
def parse_klh_sheet(sh):
    rows = sh['rows']
    out = []
    title = ' '.join(flat(v) for r in rows[:1] for v in r if v)[:300]
    hdr = next((i for i, r in enumerate(rows) if any(re.match(r'^始发站', flat(v)) for v in r)), None)
    if hdr is None:
        return out
    cells = [flat(v) for v in rows[hdr]]
    col = {k: next((j for j, c in enumerate(cells) if c.startswith(k)), None)
           for k in ('始发站', '目的站', 'SOC', 'COC', '班期', '备注', '特别')}
    hcur = {k: ('RMB' if col[k] is not None and 'RMB' in cells[col[k]] else None) for k in ('SOC', 'COC')}
    origin, sched = '', ''
    for i in range(hdr + 1, len(rows)):
        r = list(rows[i]) + [None] * 12
        if s(r[col['始发站']]):
            origin = flat(r[col['始发站']])
        if col['班期'] is not None and s(r[col['班期']]):
            sched = flat(r[col['班期']])
        dest = flat(r[col['目的站']])
        dst = dest_of(dest)
        if not dst:
            continue
        org = canon_origin(origin)
        if not org:
            continue
        remark = ' | '.join(flat(r[col[k]]) for k in ('备注', '特别') if col[k] is not None and s(r[col[k]]))[:900]
        for box in ('SOC', 'COC'):
            if col[box] is None:
                continue
            t = flat(r[col[box]])
            m = re.search(r'(USD|RMB)?\s*(\d{4,6})\s*(RMB|USD)?\s*/?\s*(40HQ|2\*20GP)?', t)
            if not t or not m:
                continue
            cur = m.group(1) or m.group(3) or hcur[box] or 'RMB'
            base = {'destination': dst['name'], 'box': box, 'size': '2x20ft' if m.group(4) == '2*20GP' else '40HQ', 'currency': cur,
                    'amount': float(m.group(2)), 'is_reference': False, 'price_note': t[m.end():].strip()[:80],
                    'origin': org, 'origin_raw': origin[:60], 'dest_raw': dest[:80],
                    'port': 'Trans-Caspian (Horgos–Aktau–Baku–Tbilisi–Halkalı)', 'schedule': sched[:160],
                    'service': ['Trans-Caspian 跨里海'], 'operator': '自营 own', 'product': 'Trans-Caspian 跨里海',
                    'remark': remark, 'header': cells[col[box]][:40], 'row': i + 1, 'col': col[box] + 1, 'title': title}
            out.append(base)
            m2 = re.search(r'SOC\s*(\d{4,6})\s*R?MB\s*/\s*2\*20GP', t)
            if m2 and box == 'SOC':
                out.append(dict(base, size='2x20ft', amount=float(m2.group(1)), price_note='per 2×20GP set',
                                col=base['col'] + 0.5))
    return out


# ---------------------------------------------------------------- main
def main():
    t0 = time.time()
    log = []
    files = []
    for dp, dn, fs in os.walk(ROOT):
        if os.path.abspath(dp).startswith(HERE):
            continue
        for f in fs:
            if not f.lower().endswith(('.xls', '.xlsx')) or f.startswith('~$') or '东南亚' in f:
                continue
            fam = next((k for k, t in FAMILIES if t(f)), None)
            if fam:
                files.append((fam, os.path.join(dp, f)))
    files.sort()
    raw, incomplete = [], False
    for fam, p in files:
        if time.time() - t0 > BUDGET:
            incomplete = True
            break
        try:
            d = load(p)
        except Exception as e:
            log.append(f'SKIP (unreadable) {p}: {e}')
            continue
        b = os.path.basename(p)
        fdate = file_date(b)
        for sh in d['sheets']:
            if '卡派' in sh['name'] or '站编' in sh['name']:
                continue
            if not any(isinstance(v, str) and DEST.search(v) for r in sh['rows'] for v in r):
                continue
            recs = parse_klh_sheet(sh) if fam == 'KLH' else parse_table_sheet(fam, sh)
            qd = (sheet_date(sh['name']) if fam == 'ZY' else None) or fdate
            for rec in recs:
                rec.update({'family': fam, 'quote_date': qd, 'file': d['path'], 'file_mtime': d['mtime'],
                            'sheet': sh['name'], 'sheet_hidden': sh['state'] != 'visible',
                            'dep_month': dep_month(b, sh['name'], rec.pop('title'))})
                raw.append(rec)
    if incomplete:
        print(f'INCOMPLETE - cached so far; re-run to continue ({time.time() - t0:.0f}s)')
        return

    # ---- latest version per (family, departure month, quote date)
    groups = {}
    for r in raw:
        groups.setdefault((r['family'], r['dep_month'], r['quote_date']), {}).setdefault(r['file'], []).append(r)
    chosen = []
    for key, byfile in sorted(groups.items(), key=lambda kv: tuple(str(x) for x in kv[0])):
        best = max(byfile, key=lambda f: ('更新' in os.path.basename(f), byfile[f][0]['file_mtime'], len(byfile[f])))
        others = sorted(set(byfile) - {best})
        log.append(f'USE  {key} -> {best}' + (f'   (ignored {len(others)} other version(s))' if others else ''))
        for r in byfile[best]:
            r['versions_ignored'] = len(others)
            chosen.append(r)

    # ---- implied RMB/USD rate per quote date, from sheets that show both
    cells = {}
    for r in chosen:
        cells.setdefault((r['file'], r['sheet'], r['row'], r['box'], r['size'], r['destination']), []).append(r)
    fx = {}
    for k, rs in cells.items():
        rmb = next((x['amount'] for x in rs if x['currency'] == 'RMB'), None)
        usd = next((x['amount'] for x in rs if x['currency'] == 'USD'), None)
        if rmb and usd and 5.5 < rmb / usd < 8 and rs[0]['quote_date']:
            fx.setdefault(rs[0]['quote_date'], []).append(rmb / usd)
    fxd = {k: round(sorted(v)[len(v) // 2], 4) for k, v in fx.items()}

    def rate_for(qd):
        if qd in fxd:
            return fxd[qd], 'same-day sheet rate'
        if not fxd or not qd:
            return 6.7, 'default 6.7'
        near = min(fxd, key=lambda k: abs((datetime.date.fromisoformat(k) - datetime.date.fromisoformat(qd)).days))
        return fxd[near], f'rate from {near} sheets'

    # ---- one record per price cell (RMB settlement + USD reference collapsed)
    out = []
    for k, rs in cells.items():
        base = next((x for x in rs if not x['is_reference'] and x['currency'] == 'RMB'), None) or \
               next((x for x in rs if not x['is_reference']), rs[0])
        rec = {x: base[x] for x in ('destination', 'quote_date', 'dep_month', 'origin', 'origin_raw', 'dest_raw', 'box', 'size',
                                    'operator', 'product', 'service', 'port', 'schedule', 'remark', 'price_note',
                                    'family', 'file', 'sheet', 'sheet_hidden', 'row', 'versions_ignored')}
        rec['family_label'] = FAMILY_LABEL[rec['family']]
        rec['native_currency'] = base['currency']
        rec['native_amount'] = round(base['amount'], 2)
        rt, why = rate_for(rec['quote_date'])
        usd_cell = next((x for x in rs if x['currency'] == 'USD'), None)
        if base['currency'] == 'USD':
            rec['usd'], rec['usd_basis'] = round(base['amount'], 2), 'quoted in USD'
        elif usd_cell:
            rec['usd'], rec['usd_basis'] = round(usd_cell['amount'], 2), 'USD reference in sheet'
        else:
            rec['usd'], rec['usd_basis'] = round(base['amount'] / rt, 2), f'converted @{rt} ({why})'
        rec['rmb'] = round(base['amount'], 2) if base['currency'] == 'RMB' else round(base['amount'] * rt, 2)
        rec['fx'] = rt
        ad = adders(rec['remark'] + ' | ' + rec['price_note'])
        rec['adders'] = ad
        excl = [a for a in ad if a.get('excluded')]
        cond = [a for a in ad if not a.get('excluded') and '绕区' in a['ctx']]
        parts, add_usd = [], 0.0
        for a in excl:
            add_usd += a['amount'] / (rt if a['cur'] == 'RMB' else 1)
            parts.append(f"{a['ctx']} +{a['amount']:.0f} {a['cur']}")
        rec['coc_pickup'] = None
        if rec['box'] == 'COC':
            pk = coc_pickup(rec['remark'], next(d['cn'] for d in DESTINATIONS if d['name'] == rec['destination']))
            city = ORIGIN_CITY.get(rec['origin'])
            if pk and city:
                hit = (0.0, 'RMB') if any(city in c for c in pk['base']) else \
                    next((v for c, v in pk['adders'].items() if city in c), (None, None))
                amt, cur = hit
                rec['coc_pickup'] = {'city': city, 'amount': amt, 'cur': cur, 'clause': pk['text']}
                if amt:
                    add_usd += amt / (rt if cur == 'RMB' else 1)
                    parts.append(f"COC box picked up in {city} +{amt:.0f} {cur}")
                elif amt == 0.0:
                    parts.append(f"COC box pickup in {city} at base price")
        rec['cond_usd'] = round(sum(a['amount'] / (rt if a['cur'] == 'RMB' else 1) for a in cond), 2) or None
        rec['cond_basis'] = ', '.join(f"{a['ctx']} +{a['amount']:.0f} {a['cur']}" for a in cond)
        rec['all_in_usd'] = round(rec['usd'] + add_usd, 2)
        rec['all_in_basis'] = '; '.join(parts) if parts else 'no extra fees listed in sheet'
        rec['route'] = f"{rec['origin']} → {rec['destination']}"
        rec['lane_id'] = ' | '.join([rec['origin'], rec['destination'], rec['operator'], rec['product'], ' / '.join(rec['service']) or 'Standard'])
        rec['source_cell'] = f"{rec['sheet']}!R{rec['row']}"
        out.append(rec)
    out.sort(key=lambda r: (r['quote_date'] or '', r['origin'], r['box'], r['size']))

    json.dump({'generated': datetime.datetime.now().isoformat(timespec='seconds'), 'fx': fxd, 'rows': out},
              open(os.path.join(HERE, 'rail_prices.json'), 'w', encoding='utf-8'), ensure_ascii=False)
    cols = ['quote_date', 'dep_month', 'route', 'lane_id', 'origin', 'destination', 'box', 'size', 'operator', 'product', 'service',
            'native_currency', 'native_amount', 'usd', 'usd_basis', 'rmb', 'all_in_usd', 'all_in_basis', 'cond_usd', 'cond_basis', 'schedule',
            'port', 'dest_raw', 'price_note', 'remark', 'family_label', 'file', 'sheet', 'source_cell',
            'sheet_hidden', 'versions_ignored']
    with open(os.path.join(HERE, 'rail_prices.csv'), 'w', encoding='utf-8-sig', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction='ignore')
        w.writeheader()
        for r in out:
            w.writerow({**r, 'service': ' / '.join(r['service'])})
    open(os.path.join(HERE, 'extract_log.txt'), 'w', encoding='utf-8').write('\n'.join(log))
    print(f'files scanned {len(files)}, prices {len(out)}, quote dates {len({r["quote_date"] for r in out})}, '
          f'{time.time() - t0:.0f}s')


if __name__ == '__main__':
    main()

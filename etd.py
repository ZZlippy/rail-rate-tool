import json,re,datetime,calendar,collections
D=json.load(open('data2.json'))
CN={'Budapest':'布达佩斯'}
PLACES=r'马拉|罗兹|杜堡|汉堡|布达佩斯|米兰|蒂尔堡|两堡|杜伊斯堡|布拉格|华沙|芬洛|安特卫普|贝尔格莱德|鹿特丹'
WD={'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'日':7,'天':7}
DATE=re.compile(r'(?<![\d.])(\d{1,2})[./](\d{1,2})(?![\d])((?:\s*[/、,，]\s*\d{1,2}(?![.\d]))*)')
WEEK=re.compile(r'(?:每)?周((?:[一二三四五六日天1-7]|[/、，,]\s*(?:每)?周?)+)')
def months_of(m):
    n=re.findall(r'\d+',m or '')
    if not n: return []
    a=int(n[0]); b=int(n[1]) if len(n)>1 else a
    return list(range(a,b+1))
def parse(sch,ms,cn,year=2026):
    t=sch or ''
    toks=[]
    for m in DATE.finditer(t):
        mo=int(m.group(1)); d1=int(m.group(2))
        if not(1<=mo<=12 and 1<=d1<=31): continue
        ds=[d1]+[int(x) for x in re.findall(r'\d{1,2}',m.group(3))]
        after=t[m.end():m.end()+2]
        if after.startswith('后'): continue
        toks.append(('d',m.start(),m.end(),mo,ds))
    for m in WEEK.finditer(t):
        body=m.group(1); days=set()
        for ch in body:
            if ch in WD: days.add(WD[ch])
            elif ch.isdigit() and 1<=int(ch)<=7: days.add(int(ch))
        if days: toks.append(('w',m.start(),m.end(),None,sorted(days)))
    toks.sort(key=lambda x:x[1])
    labs=[]
    for i,tk in enumerate(toks):
        nxt=toks[i+1][1] if i+1<len(toks) else len(t)
        labs.append(t[tk[2]:nxt])
    placed=[bool(re.search(PLACES,l)) for l in labs]
    if any(placed):
        keep=[tk for tk,l in zip(toks,labs) if cn in l]
    else:
        keep=toks
    listed=sorted({datetime.date(year,tk[3],d).isoformat() for tk in keep if tk[0]=='d' and tk[3] in ms for d in tk[4] if d<=calendar.monthrange(year,tk[3])[1]})
    wk=[tk for tk in keep if tk[0]=='w']
    weekly=[]
    if wk:
        days=set(d for tk in wk for d in tk[4])
        for mo in ms:
            for d in range(1,calendar.monthrange(year,mo)[1]+1):
                if datetime.date(year,mo,d).isoweekday() in days: weekly.append(datetime.date(year,mo,d).isoformat())
    # listed dates win when only explicit dates are given; weekly schedules are added when the sheet gives both
    if listed and weekly and any(p for p in placed):
        out={d:'weekly' for d in weekly}; out.update({d:'listed' for d in listed})
        return sorted(out.items()),'mixed'
    if listed: return [(d,'listed') for d in listed],'listed'
    if weekly: return [(d,'weekly') for d in weekly],'weekly'
    return [(datetime.date(year,mo,15).isoformat(),'assumed') for mo in ms],'assumed'
c=collections.Counter()
for r in D['rows']:
    ms=months_of(r['m'])
    e,tp=parse(r['sch'],ms,CN.get(r['ds'],r['ds']))
    r['etd']=e; r['et']=tp; c[(r['o'][:6],r['op'][:3],tp)]+=1
for k,v in sorted(c.items()): print(v,k)
# samples
for r in D['rows'][::60]: print(r['o'][:6],r['m'],r['et'],r['etd'][:4],'|',r['sch'][:70])
json.dump(D,open('data5.json','w'),ensure_ascii=False,separators=(',',':'))

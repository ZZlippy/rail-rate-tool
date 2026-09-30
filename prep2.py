import json,re
D=json.load(open('/mnt/user-data/uploads/Rail Freight/rail-price-tool/rail_prices.json',encoding='utf-8'))
CN={'Budapest':'布达佩斯','Hamburg':'汉堡','Duisburg':'杜伊斯堡'}
CITY=r'[一-龥]{2,6}'
def short(rm,dst):
    cn=CN.get(dst,dst)
    t=re.sub(r'('+CITY+r')[：:]([^：:]*?提[^：:]*?)(?=(?:'+CITY+r')[：:]|$)',lambda m: m.group(0) if m.group(1).startswith(cn[:2]) else '',rm)
    return re.sub(r'\s+',' ',t).strip()[:500]
rem=[];ri={};rows=[]
for r in D['rows']:
    s=short(r['remark'],r['destination'])
    if s not in ri: ri[s]=len(rem); rem.append(s)
    pk=r.get('coc_pickup') or {}
    rows.append({'d':r['quote_date'],'m':r['dep_month'],'o':r['origin'],'ds':r['destination'],'lid':r['lane_id'],'op':r['operator'],'pr':r['product'],
      'sv':' / '.join(r['service']) or 'Standard','bx':r['box'],'sz':r['size'],'cur':r['native_currency'],'amt':r['native_amount'],
      'usd':r['usd'],'rmb':r['rmb'],'fx':r['fx'],'ub':r['usd_basis'],'ai':r['all_in_usd'],'aib':r['all_in_basis'],
      'cu':r['cond_usd'],'cb':r['cond_basis'],'sch':r['schedule'],'port':r['port'],'dest':r['dest_raw'],'note':r['price_note'],
      'f':r['file'],'sh':r['sheet'],'r':r['row'],'rk':ri[s],'pk':pk.get('clause',''),'vi':r['versions_ignored'],'fam':r['family_label'],
      'fs':r.get('file_saved'),'cap':r.get('capacity_note',''),'so':r.get('sold_out',[])})
s=json.dumps({'generated':D['generated'],'rows':rows,'remarks':rem},ensure_ascii=False,separators=(',',':'))
open('data2.json','w').write(s); print(len(s.encode()),len(rows))

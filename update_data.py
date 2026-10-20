#!/usr/bin/env python3
"""매일 00:01 KST(=15:01 UTC)에 실행되어 data/data.json 과 index.html 내장 데이터를 갱신한다.
- 상장사(SEC 제출기업)는 EDGAR XBRL companyfacts API 로 최신 연간(10-K) 매출·영업이익·영업현금흐름·Capex 를 갱신
- 그 외 비상장/비공시 항목은 수동·리서치 입력(data/data.json)을 그대로 유지 (비공개 재무를 추정해 채우지 않음)
환경변수 SEC_UA 에 연락처 포함 User-Agent 를 지정하세요. 예) "YourOrg you@example.com"
"""
import json, os, re, sys, urllib.request, datetime, pathlib
ROOT = pathlib.Path(__file__).resolve().parent
CIK = {'USM001':90896,'USM002':278166,'USM024':752714,'USM025':1647088,'USM036':920760,'USM037':882184,'USM038':822416,'USM039':906163,'USM040':795266}
TAGS = {
 'rev':['Revenues','RevenueFromContractWithCustomerExcludingAssessedTax','SalesRevenueNet'],
 'op':['OperatingIncomeLoss'],
 'cfo':['NetCashProvidedByUsedInOperatingActivities'],
 'capex':['PaymentsToAcquirePropertyPlantAndEquipment'],
}
def get(url):
    req = urllib.request.Request(url, headers={'User-Agent': os.environ.get('SEC_UA','modular-dashboard admin@example.com')})
    with urllib.request.urlopen(req, timeout=30) as r: return json.load(r)
def latest_annual(facts, names):
    best=None
    for n in names:
        u = facts.get('us-gaap',{}).get(n,{}).get('units',{}).get('USD',[])
        for x in u:
            if x.get('form')!='10-K' or x.get('fp')!='FY' or not x.get('start'): continue
            days=(datetime.date.fromisoformat(x['end'])-datetime.date.fromisoformat(x['start'])).days
            if not 350<=days<=380: continue
            if best is None or x['end']>best['end'] or (x['end']==best['end'] and x['filed']>best['filed']): best=x
        if best: return best
    return None
def main():
    p = ROOT/'data/data.json'; d = json.load(open(p, encoding='utf-8')); changed=0
    for c in d['companies']:
        cik = CIK.get(c['id'])
        if not cik: continue
        try: facts = get(f'https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json')['facts']
        except Exception as e: print('skip', c['id'], e); continue
        vals = {k: latest_annual(facts, v) for k,v in TAGS.items()}
        if not vals['rev']: continue
        f = c.get('fin') or {}
        end = vals['rev']['end']
        # 모회사 연결 기준(McGrath, WillScot)도 동일하게 갱신. 단, 이미 더 최신 기간이면 유지.
        if f.get('periodEnd') and f['periodEnd'] > end: continue
        f.update(rev=round(vals['rev']['val']/1e6,1), periodEnd=end, conf='v',
                 src=f'https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK={cik}&type=10-K')
        f['fy'] = f"FY{end[:4]} (~{end})"
        for k in ('op','cfo','capex'):
            if vals[k]: f[k] = round(vals[k]['val']/1e6,1)
        if f.get('op') is not None and f.get('opLabel','영업이익')=='영업이익': f['margin']=round(f['op']/f['rev']*100,1)
        c['fin']=f; changed+=1
    d['generatedAt'] = datetime.datetime.now(datetime.timezone.utc).isoformat()
    json.dump(d, open(p,'w',encoding='utf-8'), ensure_ascii=False, indent=1)
    h = ROOT/'index.html'; t = h.read_text(encoding='utf-8')
    t = re.sub(r'/\*DATA_START\*/.*?/\*DATA_END\*/', lambda m: '/*DATA_START*/window.__DATA__ = '+json.dumps(d,ensure_ascii=False,separators=(',',':'))+';/*DATA_END*/', t, flags=re.S)
    h.write_text(t, encoding='utf-8'); print('updated', changed, 'companies')
if __name__=='__main__': main()

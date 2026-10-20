#!/usr/bin/env python3
"""[1단계] SEC EDGAR XBRL(무료)로 상장사/SEC 보고기업의 연간 재무를 갱신한다.

원칙
- 더 최신 회계기간(periodEnd)이 나왔을 때만 값을 교체한다. 같은 기간이면 '비어 있는 칸'만 채운다.
- 영업이익 라벨이 '영업이익' 또는 '세전이익(대용)'인 기업만 자동 갱신한다(그 외는 수동 산출값이므로 건드리지 않음).
- 환경변수 SEC_UA 에 "이름 이메일" 형태의 User-Agent 를 지정해야 한다(SEC 요구사항).
"""
import json, os, re, datetime, pathlib, urllib.request

ROOT = pathlib.Path(__file__).resolve().parent
CIK = {'USM001':90896,'USM002':278166,'USM010':1816937,'USM024':752714,'USM025':1647088,
       'USM036':920760,'USM037':882184,'USM038':822416,'USM039':906163,'USM040':795266}
TAGS = {
 'rev':['Revenues','RevenueFromContractWithCustomerExcludingAssessedTax','SalesRevenueNet'],
 'op':['OperatingIncomeLoss'],
 'pretax':['IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest',
           'IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments'],
 'cfo':['NetCashProvidedByUsedInOperatingActivities'],
 'capex':['PaymentsToAcquirePropertyPlantAndEquipment'],
}

def http_json(url, ua=None):
    req = urllib.request.Request(url, headers={'User-Agent': ua or os.environ.get('SEC_UA','modular-dashboard admin@example.com')})
    with urllib.request.urlopen(req, timeout=30) as r: return json.load(r)

def latest_annual(facts, names):
    best = None
    for n in names:
        for x in facts.get('us-gaap', {}).get(n, {}).get('units', {}).get('USD', []):
            if x.get('form') not in ('10-K','10-K/A') or x.get('fp') != 'FY' or not x.get('start'): continue
            days = (datetime.date.fromisoformat(x['end']) - datetime.date.fromisoformat(x['start'])).days
            if not 350 <= days <= 380: continue
            if best is None or x['end'] > best['end'] or (x['end'] == best['end'] and x['filed'] > best['filed']): best = x
        if best: return best
    return None

def apply_financials(d, getter=http_json, log=None):
    log = log if log is not None else []
    now = datetime.datetime.now(datetime.timezone.utc).isoformat()
    for c in d['companies']:
        cik = CIK.get(c['id'])
        if not cik: continue
        f = c.get('fin') or {}
        label = f.get('opLabel', '영업이익')
        if label not in ('영업이익', '세전이익(대용)'):
            log.append(f"{c['id']} 재무 자동갱신 제외(수동 산출 항목: {label})"); continue
        try: facts = getter(f'https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json')['facts']
        except Exception as e: log.append(f"{c['id']} SEC 재무 조회 실패: {e}"); continue
        v = {k: latest_annual(facts, t) for k, t in TAGS.items()}
        if not v['rev']: log.append(f"{c['id']} 매출 태그 없음"); continue
        end = v['rev']['end']; old_end = f.get('periodEnd')
        opk = 'pretax' if label == '세전이익(대용)' else 'op'
        new = {'rev': v['rev'], 'op': v[opk], 'cfo': v['cfo'], 'capex': v['capex']}
        if old_end and end < old_end: continue
        newer = (not old_end) or end > old_end
        if not newer and not old_end: newer = True
        changed = []
        if newer and old_end != end:
            for k, x in new.items():
                nv = round(x['val']/1e6, 1) if x else None
                if f.get(k) != nv: f[k] = nv; changed.append(k)
            f['fy'] = f"FY{end[:4]} (~{end})"
        else:  # 같은 기간: 빈 칸만 채움
            for k, x in new.items():
                if f.get(k) is None and x: f[k] = round(x['val']/1e6, 1); changed.append(k)
        if not changed: continue
        f.update(periodEnd=end, conf='v', opLabel=label, auto={'src':'SEC EDGAR XBRL','at':now})
        f['src'] = f'https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK={cik}&type=10-K'
        f['margin'] = round(f['op']/f['rev']*100, 1) if f.get('op') is not None and f.get('rev') else None
        c['fin'] = f; log.append(f"{c['id']} 재무 갱신({end}): {','.join(changed)}")
    return log

def load():
    return json.load(open(ROOT/'data/data.json', encoding='utf-8'))

def save(d):
    d['generatedAt'] = datetime.datetime.now(datetime.timezone.utc).isoformat()
    json.dump(d, open(ROOT/'data/data.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    h = ROOT/'index.html'
    if h.exists():
        t = h.read_text(encoding='utf-8')
        t = re.sub(r'/\*DATA_START\*/.*?/\*DATA_END\*/', lambda m: '/*DATA_START*/window.__DATA__ = ' + json.dumps(d, ensure_ascii=False, separators=(',', ':')) + ';/*DATA_END*/', t, flags=re.S)
        h.write_text(t, encoding='utf-8')

def main():
    d = load(); log = apply_financials(d)
    save(d); print('\n'.join(log) or '변경 없음')

if __name__ == '__main__': main()

#!/usr/bin/env python3
"""[2단계] 무료 공개 출처에서 직원·시설·프로젝트·동향을 수집해 data/data.json 에 병합한다. (유료 API·AI 호출 없음)

출처
 1) SEC EDGAR 10-K 본문  - 상장사 + SEC 보고기업(Boxabl): 직원 수, 생산시설 수
 2) 공식 사이트(robots.txt 준수) - 프로젝트/포트폴리오 목록, 직원·면적 문구, 사이트 생존 여부, RSS
 3) GDELT DOC API / Google News RSS (무료, 키 불필요) - 최근 보도, 운영중단·인수·파산 등 상태 변화 신호
 4) Wikidata (무료) - 직원 수 보조 출처(공시가 없는 비상장사, 신뢰도 '추정')

병합 원칙
 - 모든 자동값에 출처 URL, 수집일, auto 표시를 붙인다. 출처 없는 값은 저장하지 않는다.
 - 수동 입력(공시확인 'v')값은 SEC 공시의 더 새로운 수치가 아니면 덮어쓰지 않는다.
 - 프로젝트는 '자동수집(미검증)' 단계로 추가만 하고 삭제하지 않는다. 뉴스는 120일 보관.
 - 이상값(기존 대비 0.5~2배 범위 밖)은 반영하지 않고 로그에 남긴다.
"""
import argparse, datetime as dt, html, json, os, re, sys, time, urllib.error, urllib.parse, urllib.request, urllib.robotparser
import xml.etree.ElementTree as ET
from html.parser import HTMLParser
import pathlib
import update_data as U

ROOT = pathlib.Path(__file__).resolve().parent
UA = os.environ.get('CRAWL_UA', 'ModularDashboardBot/1.0 (research dashboard; contact: set CRAWL_UA env)')
SEC_UA = os.environ.get('SEC_UA', 'modular-dashboard admin@example.com')
TODAY = dt.date.today().isoformat()

# ------------------------------------------------------------------ HTTP
class Http:
    """예의 있는 HTTP 클라이언트: 호스트별 지연, robots.txt, 타임아웃. fetcher 주입으로 오프라인 테스트 가능."""
    def __init__(s, fetcher=None, delay=1.0):
        s.fetcher, s.delay, s.last, s.robots = fetcher, delay, {}, {}
    def _raw(s, url, ua, timeout=25, maxbytes=6_000_000):
        if s.fetcher: return s.fetcher(url)
        req = urllib.request.Request(url, headers={'User-Agent': ua, 'Accept': '*/*'})
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.status, r.read(maxbytes).decode(r.headers.get_content_charset() or 'utf-8', 'replace')
        except urllib.error.HTTPError as e: return e.code, ''
        except Exception: return None, ''
    def _wait(s, host, delay):
        w = s.last.get(host, 0) + delay - time.time()
        if w > 0 and not s.fetcher: time.sleep(w)
        s.last[host] = time.time()
    def allowed(s, url, ua=UA):
        u = urllib.parse.urlparse(url); base = f'{u.scheme}://{u.netloc}'
        if base not in s.robots:
            code, txt = s._raw(base + '/robots.txt', ua)
            rp = urllib.robotparser.RobotFileParser()
            if code == 200: rp.parse(txt.splitlines()); s.robots[base] = rp
            elif code in (401, 403): s.robots[base] = False          # 접근 금지로 간주
            elif code and code < 500: s.robots[base] = True            # robots 없음 → 허용
            else: s.robots[base] = False                              # 네트워크 오류 → 보수적으로 불허
        rp = s.robots[base]
        return rp if isinstance(rp, bool) else rp.can_fetch(ua, url)
    def get(s, url, ua=UA, robots=True, delay=None):
        host = urllib.parse.urlparse(url).netloc
        if robots and not s.allowed(url, ua): return 'robots', ''
        s._wait(host, s.delay if delay is None else delay)
        return s._raw(url, ua)
    def json(s, url, ua=UA, delay=None):
        code, txt = s.get(url, ua, robots=False, delay=delay)
        if code != 200: return None
        try: return json.loads(txt)
        except Exception: return None

# ------------------------------------------------------------------ HTML 파서
class Page(HTMLParser):
    SKIP = {'script', 'style', 'noscript', 'svg'}
    def __init__(s):
        super().__init__(convert_charrefs=True)
        s.tokens, s.links, s.feeds, s.title = [], [], [], ''
        s._skip, s._a, s._t = 0, None, False
    def handle_starttag(s, tag, attrs):
        a = dict(attrs)
        if tag in s.SKIP: s._skip += 1
        if tag == 'title': s._t = True
        if tag == 'a' and a.get('href'): s._a = [a['href'], '', len(s.tokens)]
        if tag == 'link' and 'alternate' in (a.get('rel') or '') and any(k in (a.get('type') or '') for k in ('rss', 'atom')):
            s.feeds.append(a.get('href'))
    def handle_endtag(s, tag):
        if tag in s.SKIP and s._skip: s._skip -= 1
        if tag == 'title': s._t = False
        if tag == 'a' and s._a:
            s.links.append((s._a[0], re.sub(r'\s+', ' ', s._a[1]).strip(), s._a[2])); s._a = None
    def handle_data(s, data):
        if s._skip: return
        t = re.sub(r'\s+', ' ', data).strip()
        if not t: return
        if s._t: s.title += t
        s.tokens.append(t)
        if s._a: s._a[1] += ' ' + t
def parse(htm):
    p = Page(); p.feed(htm); return p

# ------------------------------------------------------------------ 공통 도우미
NUM = re.compile(r'\d[\d,]*')
def to_int(s):
    m = NUM.search(str(s or '')); return int(m.group().replace(',', '')) if m else None
def norm(s): return re.sub(r'[^a-z0-9]', '', (s or '').lower().replace('’', "'"))
def ok_ratio(new, old): return old is None or old == 0 or 0.5 <= new / old <= 2.0
def title_from_slug(slug): return ' '.join(w.capitalize() for w in re.split(r'[-_]+', slug) if w)
EMP_RE = re.compile(r'(?:approximately|about|over|more than|nearly|had|employed|employs|total of)?\s*([\d,]{2,7})\+?\s+(?:full[- ]time\s+)?(?:employees|team members|associates|craftsmen|craftspeople|staff members|people)', re.I)
SQFT_RE = re.compile(r'([\d,]{4,9})[- ]?(?:sq\.?\s?ft\.?|square[- ]f(?:oo|ee)t)', re.I)
FAC_RE = re.compile(r'(\d{1,3})\s+(?:homebuilding\s+)?(?:manufacturing|production)\s+(?:facilities|plants|lines|locations)', re.I)
UNIT_RE = re.compile(r'(\d[\d,]*)\s*(units|keys|apartments|homes|beds|rooms|suites)', re.I)
STORY_RE = re.compile(r'(\d+)[- ]stor(?:y|ies)', re.I)
PROJ_PATH = re.compile(r'/(?:projects?|portfolio|our-work|work|case-stud(?:y|ies))/([^/?#]+)/?$', re.I)
GENERIC = {'click here', 'learn more', 'read more', 'view', 'view project', 'see more', 'view all', 'more', ''}
STATES = set('AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY DC AB BC ON'.split())

# ------------------------------------------------------------------ 1) SEC 10-K 본문
def strip_html(htm):
    t = re.sub(r'(?is)<(script|style).*?</\1>', ' ', htm)
    t = re.sub(r'(?s)<[^>]+>', ' ', t)
    return re.sub(r'\s+', ' ', html.unescape(t))

def sec_filing_text(http, cik):
    sub = http.json(f'https://data.sec.gov/submissions/CIK{cik:010d}.json', ua=SEC_UA, delay=0.2)
    if not sub: return None
    rec = sub.get('filings', {}).get('recent', {})
    for i, form in enumerate(rec.get('form', [])):
        if form in ('10-K', '10-K405'):
            acc = rec['accessionNumber'][i].replace('-', ''); doc = rec['primaryDocument'][i]
            url = f'https://www.sec.gov/Archives/edgar/data/{cik}/{acc}/{doc}'
            code, htm = http.get(url, ua=SEC_UA, robots=False, delay=0.2)
            if code != 200: return None
            return dict(url=url, text=strip_html(htm), report=rec.get('reportDate', [''])[i], filed=rec.get('filingDate', [''])[i])
    return None

def sec_extract(text):
    out = {}
    i = max(text.find('Human Capital'), text.find('HUMAN CAPITAL'))
    for chunk in ([text[i:i+6000]] if i >= 0 else []) + [text]:
        m = EMP_RE.search(chunk)
        if m and 50 <= int(m.group(1).replace(',', '')) <= 2_000_000:
            out['emp'] = (int(m.group(1).replace(',', '')), chunk[max(0, m.start()-40):m.end()+40]); break
    m = FAC_RE.search(text)
    if m: out['fac'] = (int(m.group(1)), text[max(0, m.start()-40):m.end()+40])
    return out

def apply_sec(c, http, log):
    from update_data import CIK
    cik = CIK.get(c['id'])
    if not cik: return False
    fl = sec_filing_text(http, cik)
    if not fl: log['errors'].append(f"{c['id']} SEC 10-K 본문 조회 실패"); return True
    ex = sec_extract(fl['text']); at = dt.datetime.now(dt.timezone.utc).isoformat()
    if 'emp' in ex:
        n, snip = ex['emp']; old = to_int((c.get('emp') or {}).get('v'))
        if not ok_ratio(n, old): log['errors'].append(f"{c['id']} 직원 수 이상값 보류: {n} (기존 {old})")
        elif old is None or abs(n-old)/max(old,1) > 0.01:
            c['emp'] = dict(v=f'{n:,}', conf='v', note=f"{fl['report'] or fl['filed']} 10-K: …{snip.strip()}…", src=fl['url'], auto={'src':'SEC 10-K','at':at}, asOf=fl['report'])
            log['changes'].append(f"{c['id']} 직원 {old}→{n}")
    if 'fac' in ex:
        n, snip = ex['fac']; fac = c.setdefault('fac', {}); old = to_int(fac.get('count'))
        if not ok_ratio(n, old): log['errors'].append(f"{c['id']} 시설 수 이상값 보류: {n} (기존 {old})")
        elif old != n:
            fac.update(count=str(n), conf='v', src=fl['url'], auto={'src':'SEC 10-K','at':at}, asOf=fl['report'])
            fac['summary'] = f"{fl['report'] or fl['filed']} 10-K: …{snip.strip()}…" + (' (이전 요약: ' + fac['summary'][:80] + ')' if fac.get('summary') else '')
            log['changes'].append(f"{c['id']} 시설 {old}→{n}")
    return True

# ------------------------------------------------------------------ 2) 공식 사이트
def same_host(u, base): return urllib.parse.urlparse(u).netloc.replace('www.', '') == urllib.parse.urlparse(base).netloc.replace('www.', '')

def analyze_site(http, base):
    res = dict(status='ok', emp=None, sqft=None, projects=[], feed=[], pages=0)
    code, htm = http.get(base)
    if code == 'robots': res['status'] = 'robots'; return res
    if code is None or code >= 400 or not htm: res['status'] = 'down'; return res
    home = parse(htm); pages = [(base, home)]
    cand = []
    for href, txt, _ in home.links:
        u = urllib.parse.urljoin(base, href.split('#')[0])
        if same_host(u, base) and re.search(r'/(about|company|team|facilit|project|portfolio|our-work|case|news|press|blog|media)', urllib.parse.urlparse(u).path, re.I) and u not in [p[0] for p in pages] and u not in cand:
            cand.append(u)
    for u in cand[:6]:
        c2, h2 = http.get(u)
        if c2 == 200 and h2: pages.append((u, parse(h2)))
    res['pages'] = len(pages)
    seen = set()
    for url, pg in pages:
        text = ' '.join(pg.tokens)
        m = EMP_RE.search(text)
        if m and not res['emp'] and 5 <= int(m.group(1).replace(',', '')) <= 500_000: res['emp'] = (int(m.group(1).replace(',', '')), url, text[max(0, m.start()-30):m.end()+30])
        m = SQFT_RE.search(text)
        if m and not res['sqft']: res['sqft'] = (m.group(1), url)
        for f in pg.feeds[:2]:
            fu = urllib.parse.urljoin(url, f or '')
            if fu not in res['feed']: res['feed'].append(fu)
        for href, txt, idx in pg.links:
            u = urllib.parse.urljoin(url, href.split('#')[0]); m = PROJ_PATH.search(urllib.parse.urlparse(u).path)
            if not m or not same_host(u, base) or 'category' in u: continue
            slug = m.group(1)
            if slug in seen: continue
            name = txt if 3 <= len(txt) <= 80 and txt.lower() not in GENERIC else ''
            if not name:  # 앵커 텍스트가 없으면 같은 링크의 다른 앵커/인접 토큰 사용 보류 → 슬러그
                name = title_from_slug(slug)
            near = ' | '.join(pg.tokens[idx:idx+8]); seen.add(slug)
            st = ''
            ms = re.search(r'-([a-z]{2})$', slug)
            if ms and ms.group(1).upper() in STATES: st = ms.group(1).upper()
            sc = []
            mu = UNIT_RE.search(near); ms2 = STORY_RE.search(near)
            if mu: sc.append(f'{mu.group(1)} {mu.group(2).lower()}')
            if ms2: sc.append(f'{ms2.group(1)} stories')
            res['projects'].append(dict(name=name, st=st, scale='; '.join(sc) or 'n/a', src=u))
    return res

def feed_items(http, url):
    code, txt = http.get(url)
    if code != 200 or not txt: return []
    try: root = ET.fromstring(txt)
    except Exception: return []
    out = []
    for it in root.iter():
        if it.tag.endswith('item') or it.tag.endswith('entry'):
            g = lambda n: next((x for x in it if x.tag.endswith(n)), None)
            t, l, d = g('title'), g('link'), g('pubDate') or g('published') or g('updated')
            link = (l.text or l.attrib.get('href')) if l is not None else ''
            if t is not None and link: out.append(dict(title=(t.text or '').strip(), url=link.strip(), date=parse_date(d.text if d is not None else ''), source='공식 피드'))
    return out[:15]

def parse_date(s):
    s = (s or '').strip()
    for f in ('%a, %d %b %Y %H:%M:%S %Z', '%a, %d %b %Y %H:%M:%S %z', '%Y-%m-%dT%H:%M:%S%z', '%Y-%m-%dT%H:%M:%SZ', '%Y%m%dT%H%M%SZ', '%Y-%m-%d'):
        try: return dt.datetime.strptime(s, f).date().isoformat()
        except Exception: pass
    return TODAY

def apply_site(c, http, log, state):
    base = c.get('site') or c.get('url')
    if not base: return
    res = analyze_site(http, base); st = state.setdefault(c['id'], {})
    if res['status'] == 'ok': st['siteFail'] = 0; st['lastOk'] = TODAY
    elif res['status'] == 'down': st['siteFail'] = st.get('siteFail', 0) + 1
    else: log['errors'].append(f"{c['id']} 사이트 수집 불가({res['status']}) {base}")
    c['alerts'] = [a for a in c.get('alerts', []) if a.get('kind') != 'site-down']
    if st.get('siteFail', 0) >= 3:
        c['alerts'].append(dict(kind='site-down', title=f"공식 사이트가 {st['siteFail']}회 연속 응답하지 않음 — 운영 상태 확인 필요", url=base, date=TODAY))
    at = dt.datetime.now(dt.timezone.utc).isoformat()
    # 프로젝트
    exist = c.setdefault('projects', []); keys = {norm(p['name']) for p in exist} | {p.get('src') for p in exist}
    added = 0
    for p in res['projects']:
        if norm(p['name']) in keys or p['src'] in keys or sum(1 for q in exist if q.get('auto')) >= 40: continue
        exist.append(dict(name=p['name'], city='', st=p['st'], market='', method='', scale=p['scale'], stage='자동수집(미검증)', note='공식 사이트 프로젝트 목록에서 자동 추출. 시기·계약 범위 확인 필요.', src=p['src'], auto=True, firstSeen=TODAY))
        keys.add(norm(p['name'])); added += 1
    if added: log['changes'].append(f"{c['id']} 프로젝트 +{added}건(사이트)")
    # 직원(공시 없는 곳만, 기존 값이 없거나 추정일 때)
    emp = c.get('emp') or {}
    if res['emp'] and emp.get('conf') != 'v':
        n, url, snip = res['emp']
        if ok_ratio(n, to_int(emp.get('v'))) and n != to_int(emp.get('v')):
            c['emp'] = dict(v=f'약 {n:,}', conf='e', note=f'공식 사이트 문구: …{snip.strip()}…', src=url, auto={'src':'공식 사이트','at':at})
            log['changes'].append(f"{c['id']} 직원 추정 {n} (사이트)")
    # 시설 면적 메모
    if res['sqft'] and c.get('fac') is not None and 'auto' not in c['fac'].get('sqftNote', ''):
        c['fac']['sqftNote'] = f"사이트 문구상 {res['sqft'][0]} sq ft (auto) — {res['sqft'][1]}"
    return res['feed']

# ------------------------------------------------------------------ 3) 뉴스
TAGS_RE = {
 'project': re.compile(r'\b(units?|apartments?|keys|groundbreaking|breaks? ground|awarded|contract|topping|delivers?|completes?|housing project|workforce housing)\b', re.I),
 'facility': re.compile(r'\b(factory|plant|facility|expan(?:d|ds|sion)|opens?|new site)\b', re.I),
 'status': re.compile(r'(bankrupt|chapter 11|insolven|receivership|foreclos|ceas(?:e|es|ed) operations|shut(?:s|ting)? down|closing (?:its |the )?(?:plant|factory|facility)|layoffs?|WARN notice|acquired by|to acquire|acquisition of|merger)', re.I),
}
KEY_REQ = re.compile(r'(modular|off-?site|prefab|factory-built|manufactured home)', re.I)
SUFFIX = re.compile(r',?\s+(Inc\.?|LLC|Corp(?:oration)?\.?|Co\.?|Ltd\.?|Company|Holdings)$', re.I)

def aliases(c):
    out = []
    for raw in [c['name'], c.get('legal') or '']:
        for part in re.split(r'\s*[/;]\s*', raw):
            part = SUFFIX.sub('', part.strip())
            if len(part) >= 4 and part.lower() not in [o.lower() for o in out]: out.append(part)
    return out[:3]

def fetch_news(http, c):
    al = aliases(c)
    if not al: return []
    q = '"%s" (modular OR prefab OR factory OR housing OR apartments)' % al[0]
    d = http.json('https://api.gdeltproject.org/api/v2/doc/doc?' + urllib.parse.urlencode(dict(query=q + ' sourcelang:english', mode='artlist', format='json', maxrecords=25, timespan='30d', sort='datedesc')), delay=5.5)
    items = [dict(title=a.get('title', ''), url=a.get('url', ''), date=parse_date(a.get('seendate')), source=a.get('domain', '')) for a in (d or {}).get('articles', [])]
    if not items:
        code, txt = http.get('https://news.google.com/rss/search?' + urllib.parse.urlencode(dict(q=al[0] + ' modular when:30d', hl='en-US', gl='US', ceid='US:en')), robots=False, delay=2)
        if code == 200:
            try:
                for it in ET.fromstring(txt).iter('item'):
                    t = (it.findtext('title') or '').strip(); src = it.findtext('source') or ''
                    items.append(dict(title=t.rsplit(' - ', 1)[0], url=(it.findtext('link') or '').strip(), date=parse_date(it.findtext('pubDate')), source=src))
            except Exception: pass
    return items

def apply_news(c, items, log):
    al = [a.lower() for a in aliases(c)]; need_kw = c['primary'] in ('builder', 'developer', 'gc')
    keep = []
    for it in items:
        t = it['title']
        if not it.get('url') or not any(a in t.lower() for a in al): continue
        if need_kw and not KEY_REQ.search(t) and it.get('source') != '공식 피드': continue
        it['tags'] = [k for k, rx in TAGS_RE.items() if rx.search(t)]
        keep.append(it)
    old = {n['url']: n for n in c.get('news', [])}
    for it in keep:
        if it['url'] not in old: old[it['url']] = it; log['changes'].append(f"{c['id']} 뉴스+ {it['title'][:50]}")
    cut = (dt.date.today() - dt.timedelta(days=120)).isoformat()
    c['news'] = sorted([n for n in old.values() if n.get('date', TODAY) >= cut], key=lambda n: n.get('date', ''), reverse=True)[:15]
    c['alerts'] = [a for a in c.get('alerts', []) if a.get('kind') == 'site-down']
    c['alerts'] += [dict(kind='news-status', title=n['title'], url=n['url'], date=n['date']) for n in c['news'] if 'status' in n.get('tags', []) and n.get('date', '') >= (dt.date.today() - dt.timedelta(days=60)).isoformat()][:3]

# ------------------------------------------------------------------ 4) Wikidata (보조)
def wikidata_emp(http, c):
    for name in aliases(c)[:1]:
        s = http.json('https://www.wikidata.org/w/api.php?' + urllib.parse.urlencode(dict(action='wbsearchentities', search=name, language='en', format='json', limit=3)), delay=1)
        for r in (s or {}).get('search', []):
            if r.get('label', '').lower() == name.lower() and re.search(r'company|manufactur|builder|construction|housing|developer|real estate|homebuild', r.get('description', ''), re.I):
                e = http.json(f"https://www.wikidata.org/wiki/Special:EntityData/{r['id']}.json", delay=1)
                cl = (e or {}).get('entities', {}).get(r['id'], {}).get('claims', {}).get('P1128', [])
                best = None
                for x in cl:
                    try:
                        n = int(float(x['mainsnak']['datavalue']['value']['amount'])); y = x.get('qualifiers', {}).get('P585', [{}])[0].get('datavalue', {}).get('value', {}).get('time', '')[1:5]
                        if best is None or y > best[1]: best = (n, y)
                    except Exception: pass
                if best: return best[0], best[1], f"https://www.wikidata.org/wiki/{r['id']}"
    return None

def apply_wikidata(c, http, log):
    emp = c.get('emp') or {}
    if emp.get('v'): return
    w = wikidata_emp(http, c)
    if w and 10 <= w[0] <= 2_000_000:
        c['emp'] = dict(v=f'약 {w[0]:,}' + (f' ({w[1]})' if w[1] else ''), conf='e', note='Wikidata 보조 출처 — 정확성 확인 필요', src=w[2], auto={'src':'Wikidata','at':dt.datetime.now(dt.timezone.utc).isoformat()})
        log['changes'].append(f"{c['id']} 직원 추정 {w[0]} (Wikidata)")

# ------------------------------------------------------------------ 실행
def run(d, http, only=None, deadline=None):
    state_p = ROOT/'data/state.json'
    state = json.load(open(state_p)) if state_p.exists() else {}
    log = dict(runAt=dt.datetime.now(dt.timezone.utc).isoformat(), changes=[], errors=[])
    for line in U.apply_financials(d):   # 1단계 재무(SEC XBRL)
        (log['errors'] if any(k in line for k in ('실패', '제외', '없음')) else log['changes']).append(line)
    feeds = {}
    for c in d['companies']:
        if only and c['id'] not in only: continue
        if deadline and time.time() > deadline: log['errors'].append('시간 제한으로 일부 기업 건너뜀'); break
        for step, fn in (('SEC', lambda: apply_sec(c, http, log)), ('사이트', lambda: feeds.__setitem__(c['id'], apply_site(c, http, log, state) or [])),
                         ('뉴스', lambda: apply_news(c, fetch_news(http, c) + (feed_items(http, feeds[c['id']][0]) if feeds.get(c['id']) else []), log)), ('Wikidata', lambda: apply_wikidata(c, http, log))):
            try: fn()
            except Exception as e: log['errors'].append(f"{c['id']} {step} 오류: {type(e).__name__}: {e}")
    json.dump(state, open(state_p, 'w'), ensure_ascii=False, indent=1)
    return log

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--only', nargs='*'); ap.add_argument('--max-minutes', type=int, default=40)
    a = ap.parse_args()
    d = U.load(); http = Http()
    log = run(d, http, a.only, time.time() + a.max_minutes * 60)
    U.save(d)
    lp = ROOT/'data/auto_log.json'; hist = json.load(open(lp, encoding='utf-8')) if lp.exists() else {'runs': []}
    hist['runs'] = ([dict(log, changes=log['changes'][:200], errors=log['errors'][:100])] + hist['runs'])[:30]
    json.dump(hist, open(lp, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f"변경 {len(log['changes'])}건, 오류 {len(log['errors'])}건"); [print(' -', x) for x in log['errors'][:20]]

if __name__ == '__main__': main()

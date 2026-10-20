import sys, json, pathlib, unittest, copy
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import collect as C, update_data as U

SITE = '''<html><head><title>RISE</title><link rel="alternate" type="application/rss+xml" href="/feed/"></head><body>
<a href="/about/">About us</a>
<div><h2><a href="https://risemodular.com/project/cambria-hotel-ofallon/">Cambria Hotel O'Fallon</a></h2><p>4 stories | 118 keys | St. Louis suburb</p><a href="/project/cambria-hotel-ofallon/">Click here</a></div>
<div><h2><a href="/project/dry-creek-apartments/">Dry Creek Apartments</a></h2><p>5 buildings | 4 stories each | 326 units | Denver suburb</p></div>
<div><a href="/project/new-flats-denver-co/">New Flats</a><p>12 units</p></div>
<a href="/our-work-category/hotels/">Hotels</a><a href="/project/">All projects</a>
</body></html>'''
ABOUT = '<html><body><p>We have more than 150 employees at our Owatonna facility of 120,000 square feet.</p></body></html>'
FEED = '<rss><channel><item><title>RISE Modular wins 200-unit apartment contract</title><link>https://risemodular.com/news/1</link><pubDate>Mon, 12 Oct 2026 10:00:00 +0000</pubDate></item></channel></rss>'
SUB = json.dumps({'filings': {'recent': {'form': ['8-K', '10-K'], 'accessionNumber': ['0001-26-1', '0002-26-2'], 'primaryDocument': ['a.htm', 'k.htm'], 'reportDate': ['', '2026-03-28'], 'filingDate': ['2026-06-01', '2026-05-22']}}})
def tenk(n): return f'<html><body><h2>Human Capital</h2><p>As of year end we employed approximately {n} full-time team members.</p><p>We operate 33 homebuilding production lines.</p></body></html>'
GDELT = json.dumps({'articles': [
  {'title': 'RISE Modular breaks ground on 120-unit apartment project', 'url': 'https://n.com/a', 'seendate': '20261010T120000Z', 'domain': 'n.com'},
  {'title': 'RISE Modular files Chapter 11 bankruptcy', 'url': 'https://n.com/b', 'seendate': '20261012T120000Z', 'domain': 'n.com'},
  {'title': 'Unrelated company news', 'url': 'https://n.com/c', 'seendate': '20261012T120000Z', 'domain': 'n.com'}]})
WD_S = json.dumps({'search': [{'id': 'Q1', 'label': 'Deltec Homes', 'description': 'American home manufacturer company'}]})
WD_E = json.dumps({'entities': {'Q1': {'claims': {'P1128': [{'mainsnak': {'datavalue': {'value': {'amount': '+85'}}}, 'qualifiers': {'P585': [{'datavalue': {'value': {'time': '+2023-00-00T00:00:00Z'}}}]}}]}}}})

def mk(emp_doc=None, down=False):
    pages = {'https://risemodular.com/robots.txt': (404, ''), 'https://risemodular.com/': (503, '') if down else (200, SITE),
             'https://risemodular.com/about/': (200, ABOUT), 'https://risemodular.com/feed/': (200, FEED),
             'https://data.sec.gov/submissions/CIK0000278166.json': (200, SUB),
             'https://www.sec.gov/Archives/edgar/data/278166/0002262/k.htm': (200, emp_doc or tenk('7,700')),
             'https://www.wikidata.org/w/api.php': (200, WD_S), 'https://www.wikidata.org/wiki/Special:EntityData/Q1.json': (200, WD_E)}
    def f(url):
        if url in pages: return pages[url]
        if url.startswith('https://api.gdeltproject.org'): return 200, GDELT
        if url.startswith('https://www.wikidata.org/w/api.php'): return 200, WD_S
        if url.startswith('https://news.google.com'): return 200, '<rss/>'
        return 404, ''
    return C.Http(fetcher=f)

def comp(**k):
    c = dict(id='USM008', name='RISE Modular', legal='RISE Modular', primary='gc', url='https://risemodular.com/', fin=None,
             emp={'v': None, 'conf': None}, fac={'count': '1', 'list': [], 'summary': ''}, projects=[dict(name='Cambria Hotel O’Fallon', src='x')])
    c.update(k); return c

class T(unittest.TestCase):
    def test_site_projects_dedupe_and_scale(self):
        c, log = comp(), dict(changes=[], errors=[]); feeds = C.apply_site(c, mk(), log, {})
        names = [p['name'] for p in c['projects']]
        self.assertEqual(names.count("Cambria Hotel O'Fallon") + names.count('Cambria Hotel O’Fallon'), 1)   # 곡선 따옴표 중복 제거
        self.assertIn('Dry Creek Apartments', names)
        dc = next(p for p in c['projects'] if p['name'] == 'Dry Creek Apartments'); self.assertIn('326 units', dc['scale'])
        nf = next(p for p in c['projects'] if p['name'] == 'New Flats'); self.assertEqual(nf['st'], 'CO'); self.assertTrue(nf['auto'])
        self.assertNotIn('Hotels', names)   # 카테고리 페이지 제외
        self.assertEqual(c['emp']['v'], '약 150'); self.assertEqual(c['emp']['conf'], 'e')
        self.assertTrue(feeds and feeds[0].endswith('/feed/'))
    def test_site_down_alert_after_3(self):
        c, st = comp(), {}
        for i in range(3): C.apply_site(c, mk(down=True), dict(changes=[], errors=[]), st)
        self.assertTrue(any(a['kind'] == 'site-down' for a in c['alerts']))
        C.apply_site(c, mk(), dict(changes=[], errors=[]), st); self.assertFalse(any(a['kind'] == 'site-down' for a in c['alerts']))
    def test_robots_block(self):
        h = C.Http(fetcher=lambda u: (200, 'User-agent: *\nDisallow: /') if u.endswith('robots.txt') else (200, SITE))
        self.assertEqual(C.analyze_site(h, 'https://risemodular.com/')['status'], 'robots')
    def test_sec_text(self):
        c = dict(id='USM002', emp={'v': '약 7,700', 'conf': 'v'}, fac={'count': '33 라인 (미국 31 + 멕시코 2)'}); log = dict(changes=[], errors=[])
        C.apply_sec(c, mk(), log); self.assertEqual((log['changes'], log['errors']), ([], []))                     # 동일 값 → 변경 없음
        C.apply_sec(c, mk(tenk('8,100')), log); self.assertEqual(c['emp']['v'], '8,100'); self.assertEqual(c['emp']['conf'], 'v')
        C.apply_sec(c, mk(tenk('40,000')), log); self.assertEqual(c['emp']['v'], '8,100'); self.assertTrue(any('이상값' in e for e in log['errors']))
    def test_news_alerts_and_filter(self):
        c, log = comp(), dict(changes=[], errors=[]); C.apply_news(c, C.fetch_news(mk(), c), log)
        self.assertEqual(len(c['news']), 2); self.assertEqual(c['news'][0]['title'], 'RISE Modular files Chapter 11 bankruptcy')
        self.assertTrue(any(a['kind'] == 'news-status' for a in c['alerts'])); self.assertIn('project', c['news'][1]['tags'])
    def test_wikidata(self):
        c = dict(id='USM014', name='Deltec Homes', legal='Deltec Homes', emp={'v': None}); C.apply_wikidata(c, mk(), dict(changes=[], errors=[]))
        self.assertEqual(c['emp']['conf'], 'e'); self.assertIn('85', c['emp']['v'])
    def test_xbrl_no_downgrade(self):
        facts = {'us-gaap': {'Revenues': {'units': {'USD': [{'form': '10-K', 'fp': 'FY', 'start': '2025-03-30', 'end': '2026-03-28', 'val': 2.9e9, 'filed': '2026-05-22'}]}}}}
        d = {'companies': [dict(id='USM002', fin=dict(fy='FY2026', rev=2244.5, op=228.6, cfo=None, capex=35.4, opLabel='영업이익', periodEnd='2026-03-28', conf='v'))]}
        U.apply_financials(d, getter=lambda u: {'facts': facts}); f = d['companies'][0]['fin']
        self.assertEqual(f['rev'], 2244.5)   # 같은 기간 → 기존 검증값 유지, 빈 칸만 채움
        d['companies'][0]['fin']['opLabel'] = '주택사업 영업이익'; U.apply_financials(d, getter=lambda u: {'facts': facts}); self.assertEqual(d['companies'][0]['fin']['rev'], 2244.5)
if __name__ == '__main__': unittest.main(verbosity=2)

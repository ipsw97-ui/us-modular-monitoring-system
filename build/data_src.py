# -*- coding: utf-8 -*-
"""미국 모듈러 인텔리전스 데이터 소스.
엑셀(CompanyMaster/Financials/Factories/Projects/Sources)을 기반으로 하고,
공개자료로 확인한 값은 conf='v'(공시확인), 지식기반 추정은 'e'(추정·검증필요), 없으면 None 으로 표기한다.
"""
import openpyxl, json, re, sys

XLSX = sys.argv[1] if len(sys.argv) > 1 else None

# ---- 역할(구분) 키 ----
ROLE = {
    'provider': '제조·공급(Provider)',
    'panel': '패널 공급(Panelized)',
    'developer': '개발사(Developer)',
    'builder': '빌더(Builder)',
    'gc': '종합건설(GC)',
    'lessor': '임대·리스(Relocatable)',
    'tech': '기술(Tech)',
    'dealer': '딜러·유통(Dealer)',
}

def H(city, st, lat, lon, conf='v'):
    return dict(city=city, state=st, lat=lat, lon=lon, conf=conf)

# id: 보강 데이터 (엑셀에 없는 항목)
X = {
 'USM001': dict(hq=H('Troy','MI',42.60,-83.15), roles=['provider','builder'], primary='provider',
   types=['Volumetric','Hybrid','HUD-Manufactured'], regions=['Nationwide','Midwest','South','West'],
   fin=dict(periodEnd='2026-03-28', fy='FY2026 (~2026-03-28)', rev=2663.6, op=251.8, cfo=303.9, capex=34.1, conf='v',
            src='https://www.sec.gov/Archives/edgar/data/0000090896/000119312526237809/sky-ex99_1.htm',
            note='연결 기준(HUD-code 포함). 모듈러 단독 매출은 별도 분리 필요.'),
   emp=dict(v='9,300+', conf='v', note='FY2026 실적발표 기준 "more than 9,300"'),
   fac=dict(count='46', conf='v', summary='미국 및 캐나다 서부 46개 제조시설(FY2026 8-K). 개별 공장 좌표는 미공개/미확인 → 지도 미표시.', list=[]),
   ov='Champion Homes(구 Skyline Champion)는 미국·캐나다 서부에 46개 공장을 운영하는 북미 최대급 공장제작 주택 업체. 모듈러(IRC)·HUD 매뉴팩처드·상업용 모듈러까지 포괄하며 Excel Homes 등 모듈러 브랜드를 보유.',
   port=['Single-family','Manufactured housing','Modular','상업용 모듈러'], projects=[], brands=['Excel Homes','Titan','Moduline']),
 'USM002': dict(hq=H('Phoenix','AZ',33.45,-112.07), roles=['provider','builder','dealer'], primary='provider',
   types=['Volumetric','Hybrid','HUD-Manufactured','Park Model'], regions=['Sun Belt','Southwest','South','West','Nationwide'],
   fin=dict(periodEnd='2026-03-28', fy='FY2026 (~2026-03-28)', rev=2244.5, op=228.6, cfo=None, capex=35.4, conf='v',
            src='https://www.sec.gov/Archives/edgar/data/0000278166/000162828026037575/cvco-2026328xex991.htm',
            note='영업현금흐름은 10-K 현금흐름표 추출 대기. 연결 기준(모듈러 외 매출 포함).'),
   emp=dict(v='약 7,700', conf='v', note='FY2026 10-K: full-time team members 약 7,700명'),
   fac=dict(count='33 라인 (미국 31 + 멕시코 2)', conf='v', summary='FY2026 10-K: 주택 생산라인 33개(미국 31, 멕시코 Ojinaga 2). 주별 분포는 하단 목록 참조.', list=[
     ('Austin, TX',30.27,-97.74),('Fort Worth, TX (2)',32.75,-97.33),('Lancaster, TX',32.59,-96.76),('Seguin, TX',29.57,-97.96),('Waco, TX',31.55,-97.15),
     ('Glendale, AZ',33.54,-112.19),('Goodyear, AZ',33.44,-112.36),('Phoenix, AZ',33.45,-112.07),('Nappanee, IN',41.44,-86.00),('Goshen, IN',41.58,-85.83),
     ('Millersburg, OR',44.67,-123.06),('Woodburn, OR',45.14,-122.86),('Douglas, GA',31.51,-82.85),('Moultrie, GA',31.18,-83.79),
     ('Shippenville, PA (2)',41.25,-79.40),('Emlenton, PA',41.18,-79.71),('Martinsville, VA',36.69,-79.87),('Rocky Mount, VA',36.99,-79.89),
     ('Crouse, NC',35.40,-81.30),('Hamlet, NC',34.89,-79.70),('Ocala, FL',29.19,-82.14),('Plant City, FL',28.02,-82.11),('Riverside, CA',33.95,-117.40),
     ('Nampa, ID',43.54,-116.56),('Deming, NM',32.27,-107.76),('Duncan, OK',34.50,-97.96),('Montevideo, MN',44.94,-95.72),('Dorchester, WI',44.99,-90.31),('Lafayette, TN',36.52,-86.03)]),
   ov='Cavco는 애리조나 피닉스에 본사를 둔 상장사로 제조주택·모듈러·파크모델·ADU를 생산하고, 리테일·금융(CountryPlace)·보험(Standard Casualty)까지 수직계열화. 선벨트(텍사스·애리조나·플로리다) 중심.',
   port=['Single-family','Manufactured housing','Modular','ADU','Park model'], projects=[]),
 'USM003': dict(hq=H('Maryville','TN',35.76,-83.97), roles=['provider','builder','dealer'], primary='provider',
   types=['Volumetric','Hybrid','HUD-Manufactured'], regions=['Nationwide','South','Sun Belt','Southeast'],
   fin=None, emp=dict(v='1만 명 이상 (추정)', conf='e', note='Berkshire 산하, 단독 공시 없음. 검증 필요.'),
   fac=dict(count=None, conf=None, summary='미국 내 다수 공장 보유(개수 미검증). Berkshire Hathaway 연차보고서/자사 사이트로 검증 예정.', list=[]),
   ov='Berkshire Hathaway 산하 미국 최대 규모 주택 제조·리테일·금융 통합 업체. 모듈러·HUD 주택, 임대·금융(Vanderbilt)·자체 리테일(Clayton Home Centers)을 수직 통합.',
   port=['Single-family','Manufactured housing','Modular','ADU'], projects=[], fin_note='자회사 단독 재무 미공개. Berkshire 연결 공시에 포함.'),
 'USM004': dict(hq=H('Boise','ID',43.62,-116.20), roles=['provider','gc'], primary='provider',
   types=['Volumetric'], regions=['West','Mountain West'],
   fin=None, emp=dict(v=None, conf=None, note='미공개'),
   fac=dict(count='1', conf='v', summary='Boise, ID 소재 공장·본사 (5556 Federal Way). 설비 면적·생산능력 미공개.', list=[('Boise, ID (본사·공장)',43.62,-116.20)]),
   ov='아이다호 보이시의 목조 다층 볼류메트릭 모듈러 제조사. 2001년 이래 200+ 프로젝트를 납품했고, 호텔·멀티패밀리·워크포스·학생주택에 강점. 2020년 1월 대주단 주도의 경매로 창업자 측에 매각되어 재출범.',
   port=['Hotel','Multifamily','Workforce housing','Student housing'],
   site='https://www.guerdonmodularbuildings.com/',
   projects=[dict(name='서부 12개 주 + 캐나다 납품 이력 (200+ 프로젝트)', city='-', st='West', market='Hotel / Multifamily', method='Volumetric', scale='200+ projects', stage='Completed (누적)', src='https://housinginnovation.co/factory/guerdon-modular/'), dict(name='Eviva', city='Sacramento', st='CA', market='Market-rate apartments', method='Volumetric', scale='n/a', stage='포트폴리오 게시(시기 미확인)', src='https://www.guerdonmodularbuildings.com/project/eviva-sacrament-ca/'), dict(name='Cubix', city='Seattle', st='WA', market='Micro-suite apartments', method='Volumetric', scale='n/a', stage='포트폴리오 게시(시기 미확인)', src='https://www.guerdonmodularbuildings.com/project/cubix-seattle-wa/'), dict(name='Domain', city='San Jose', st='CA', market='Apartment complex', method='Volumetric', scale='n/a', stage='포트폴리오 게시(시기 미확인)', src='https://www.guerdonmodularbuildings.com/project/domain-san-jose-ca/'), dict(name='Cheecham Lodge', city='Anzac', st='AB', market='Market-rate apartments', method='Volumetric', scale='n/a', stage='포트폴리오 게시(시기 미확인)', src='https://www.guerdonmodularbuildings.com/project/cheecham-lodge-anzac-ab/'), dict(name='Home2 Suites', city='San Francisco', st='CA', market='Hotel', method='Volumetric', scale='n/a', stage='포트폴리오 게시(시기 미확인)', src='https://www.guerdonmodularbuildings.com/project/home2-suites-sanfrancisco-ca/'), dict(name='The Union Flats', city='Union City', st='CA', market='Apartment community (4 buildings)', method='Volumetric', scale='n/a', stage='포트폴리오 게시(시기 미확인)', src='https://www.guerdonmodularbuildings.com/project/the-union-flats-union-city-ca/')],
   note='2015–2019년 미국 주요 호텔 브랜드 모듈러 호텔의 70% 이상을 시공했다고 소개(제3자 정리).'),
 'USM005': dict(hq=H('Buena Vista','CO',38.84,-106.13,'e'), roles=['provider','developer','builder'], primary='provider',
   types=['Volumetric','Hybrid'], regions=['Mountain West','West'],
   fin=None, emp=dict(v='약 130', conf='v', note='Bracewell 블로그 기준'),
   fac=dict(count='1', conf='v', summary='Buena Vista, CO 11만 sq ft 공장. 1일 1세대 생산 가능(표준 설계 최단 9일), 연 수백 세대.', list=[('Buena Vista, CO',38.84,-106.13)]),
   ov='콜로라도 산악지역 워크포스·멀티패밀리·마이크로홈을 개발·생산하는 모듈러 개발/제조 통합 업체. 병원·학교·지자체 연계 PPP 프로젝트에 특화. 최근 일부 프로젝트 분쟁 보도가 있어 채권·사업 리스크 모니터링 필요.',
   port=['Workforce housing','Multifamily','Single-family','Micro-home'],
   projects=[
     dict(name='The Farm at Buena Vista', city='Buena Vista', st='CO', market='Housing', method='Volumetric', scale='n/a', stage='Completed / 운영', src='https://www.bc.com/blog/fading-west-a-growing-force-in-workforce-housing/'),
     dict(name='Aspen Health 병원 워크포스 주택', city='Aspen 일대', st='CO', market='Workforce housing', method='Volumetric', scale='n/a', stage='진행 중', src='https://www.bc.com/blog/fading-west-a-growing-force-in-workforce-housing/'),
     dict(name='Gunnison 시 공공·민간 협력(PPP)', city='Gunnison', st='CO', market='Workforce housing', method='Volumetric', scale='n/a', stage='진행 중', src='https://www.bc.com/blog/fading-west-a-growing-force-in-workforce-housing/'),
     dict(name='Vail School District 교직원 주택 (Habitat for Humanity Vail Valley)', city='Vail', st='CO', market='Workforce housing', method='Volumetric', scale='12 homes', stage='진행 중', src='https://www.bc.com/blog/fading-west-a-growing-force-in-workforce-housing/')]),
 'USM006': dict(hq=H('New York','NY',40.71,-74.00,'l'), roles=['provider','gc'], primary='provider', types=['Volumetric'], regions=['Northeast','Mid-Atlantic'],
   ov='뉴욕 등 도심 인필 멀티패밀리·호스피탈리티용 볼류메트릭 모듈러 업체. 운영 상태·소유권·공장 현황이 불명확하여 사용 전 검증이 필요.'),
 'USM007': dict(hq=H('Pennsylvania (추정)','PA',40.27,-76.88,'l'), roles=['provider','gc'], primary='provider', types=['Volumetric','Hybrid'], regions=['Mid-Atlantic','Northeast'],
   ov='멀티패밀리·호스피탈리티·기관 건물 중심의 디자인-빌드형 모듈러 제조사. 활성 공장·최근 주거 프로젝트 검증 필요.'),
 'USM008': dict(hq=H('Minneapolis','MN',44.98,-93.27,'v'), roles=['provider','gc'], primary='gc', types=['Volumetric'], regions=['Midwest','Mountain West'],
   fac=dict(count='1', conf='v', summary='공식 사이트 푸터: 본사 Minneapolis, MN / 제조시설 Owatonna, MN. 면적·생산능력은 미공개.', list=[('Owatonna, MN (RISE Manufacturing Facility)',44.08,-93.23)]),
   ov='모듈러 제조·설치·GC 서비스를 일괄 제공하는 미네소타 기반 볼류메트릭 업체. 공식 포트폴리오에 호텔·학생주택·멀티패밀리 6건이 게시(미주리·위스콘신·콜로라도 등). 법인 현황 검증 필요.'),
 'USM009': dict(hq=H('Nampa','ID',43.54,-116.56,'e'), roles=['provider','tech'], primary='provider', types=['Volumetric'], regions=['West','Mountain West'],
   fac=dict(count='1', conf='e', summary='Nampa, ID 자동화 라인 기반 공장(추정). 공식 검증 필요.', list=[('Nampa, ID',43.54,-116.56)]),
   ov='자동화 라인 기반 볼류메트릭 모듈러 제조사(아이다호). 멀티패밀리·저렴주택 중심. 현재 수주잔고 검증 필요.'),
 'USM010': dict(hq=H('North Las Vegas','NV',36.20,-115.12,'v'), roles=['provider','tech'], primary='tech', types=['Volumetric','Foldable module'], regions=['Southwest','Nationwide'],
   emp=dict(v='128', conf='v', note='FY2025 10-K: 2025-12-31 기준 128명 (전년 167명)'),
   fac=dict(count='3', conf='v', summary='FY2025 10-K: North Las Vegas 임차 공장 3개(약 42만 sq ft). 발표 capacity와 실제 인도량(2025년 Casita 23대/15개 고객) 구분 필요.', list=[('North Las Vegas, NV – Factory 1 (5345 E. N. Belt Rd)',36.20,-115.12),('North Las Vegas, NV – Factory 2',36.20,-115.115),('North Las Vegas, NV – Factory 3',36.205,-115.12)]),
   ov='접이식(foldable) 볼류메트릭 모듈 "Casita"를 대량생산하려는 네바다 기반 업체. 비상장이지만 SEC에 10-K를 제출하는 보고 기업이라 재무·직원·시설 공시가 가능. ADU·단독주택·재난주택 타깃.',
   note='SEC 보고기업(CIK 1816937). 2025년 인도 Casita 23대로 발표 capacity 대비 실적이 작음.'),
 'USM011': dict(hq=H('California','CA',34.11,-117.29,'l'), roles=['provider','builder'], primary='provider', types=['Panelized','Volumetric'], regions=['West'],
   ov='캘리포니아 중심 프리팹 주택·ADU 업체. 현재 운영 여부·공장·프로젝트 파이프라인 검증 필요.'),
 'USM012': dict(hq=H('Modesto (추정)','CA',37.64,-120.99,'l'), roles=['panel','provider'], primary='panel', types=['Panelized'], regions=['South','Southeast','West'],
   ov='정밀 패널화(Panelized) 프레임 시스템 업체(Entekra). 소유권·공장 상태·제품라인 검증 필요 — 엑셀 상 "RISE / Entekra" 표기 정정 필요.'),
 'USM013': dict(hq=H('Wausau','WI',44.96,-89.63), roles=['panel','builder'], primary='panel', types=['Panelized'], regions=['Midwest'],
   ov='위스콘신 기반의 패널화 주택 공급 및 빌더 네트워크. 운영 모델(패널 공급/딜러망) 확인 필요.'),
 'USM014': dict(hq=H('Asheville','NC',35.60,-82.55), roles=['provider','panel'], primary='provider', types=['Panelized','Component-based','Prefabricated'], regions=['Southeast','Nationwide'],
   fac=dict(count='1', conf='e', summary='노스캐롤라이나 Asheville 일대 공장(추정).', list=[('Asheville, NC',35.60,-82.55)]),
   ov='허리케인 대응 등 내재해성 프리팹 주택 시스템에 특화된 노스캐롤라이나 업체. 전국 납품.'),
 'USM015': dict(hq=H('Pennsylvania (추정)','PA',40.04,-76.30,'l'), roles=['dealer','builder'], primary='dealer', types=['Volumetric','Manufactured/Modular mix'], regions=['Nationwide'],
   ov='전국 딜러/빌더 네트워크 기반 모듈러 주택 공급사. 직접 제조 여부 검증 필요.'),
 'USM016': dict(hq=H('Martinsville','VA',36.69,-79.87,'e'), roles=['provider'], primary='provider', types=['Modular','Manufactured'], regions=['Southeast','Mid-Atlantic'],
   fac=dict(count=None, conf='e', summary='버지니아 Martinsville 소재(추정). 현재 활성 공장 확인 필요.', list=[('Martinsville, VA',36.69,-79.87)]),
   ov='버지니아 기반 모듈러·제조주택 업체. 제품 카테고리·활성 공장 확인 필요.'),
 'USM017': dict(hq=H('Liverpool','PA',40.57,-76.98,'e'), roles=['provider'], primary='provider', types=['Modular'], regions=['Northeast','Mid-Atlantic'],
   fac=dict(count=None, conf='e', summary='펜실베이니아 Liverpool 공장(추정). Champion Homes 계열 브랜드 여부 확인 필요.', list=[('Liverpool, PA',40.57,-76.98)]),
   ov='동북부·중부대서양 단독주택 모듈러 브랜드. 현 브랜드 소유자(Champion 계열로 알려짐)와 생산 거점 검증 필요.'),
 'USM018': dict(hq=H('Pennsylvania (추정)','PA',40.50,-77.00,'l'), roles=['provider'], primary='provider', types=['Volumetric'], regions=['Northeast','Mid-Atlantic'], ov='볼류메트릭 모듈러 제조사. 공식 사이트·활성 여부 검증 필요(MBI 디렉터리 활용).'),
 'USM019': dict(hq=H('Wingdale','NY',41.64,-73.57,'e'), roles=['provider','builder'], primary='provider', types=['Volumetric'], regions=['Northeast'],
   fac=dict(count=None, conf='e', summary='뉴욕주 Wingdale 일대(추정).', list=[('Wingdale, NY',41.64,-73.57)]),
   ov='뉴욕·코네티컷 등 동북부 단독주택·ADU 모듈러 공급사. 빌더 네트워크 기반.'),
 'USM020': dict(hq=H('Scranton','PA',41.41,-75.66,'e'), roles=['provider'], primary='provider', types=['Volumetric'], regions=['Northeast','Mid-Atlantic'],
   fac=dict(count=None, conf='e', summary='펜실베이니아 Scranton 일대(추정).', list=[('Scranton, PA',41.41,-75.66)]),
   ov='동북부·중부대서양 단독주택·멀티패밀리 볼류메트릭 모듈러 제조사.'),
 'USM021': dict(hq=H('Georgia (추정)','GA',33.75,-84.39,'l'), roles=['provider'], primary='provider', types=['Volumetric'], regions=['Southeast'], ov='남동부 볼류메트릭 모듈러 제조사. MBI 디렉터리로 공식 정보 검증 필요.'),
 'USM022': dict(hq=H('Pennsylvania (추정)','PA',40.90,-77.80,'l'), roles=['provider','gc'], primary='provider', types=['Volumetric','Hybrid'], regions=['Northeast','Nationwide'], ov='캐나다 NRB 계열 디자인-빌드 모듈러. 미국 내 거점·소유구조 검증 필요.'),
 'USM023': dict(hq=H('Florida (추정)','FL',26.34,-81.78,'l'), roles=['provider','lessor'], primary='lessor', types=['Volumetric'], regions=['Nationwide'], ov='교육·의료·워크포스용 이동식(relocatable) 모듈러 제조·임대. 영구형 주거와 분리 태깅 필요.'),
 'USM024': dict(hq=H('Livermore','CA',37.68,-121.77), roles=['lessor','provider'], primary='lessor', types=['Volumetric'], regions=['Nationwide','West'],
   fin=dict(periodEnd='2025-12-31', fy='FY2025 (CY2025)', rev=944.2, op=243.6, cfo=255.7, capex=None, conf='v',
            src='https://www.sec.gov/Archives/edgar/data/752714/000119312526071461/mgrc-ex99_1.htm',
            note='모회사 McGrath RentCorp 연결 기준. Mobile Modular 부문 매출 $645.1M (총매출의 약 68%).'),
   ov='McGrath RentCorp의 이동식 모듈러 임대·판매 사업부. 교육(학교 증축)·상업·헬스케어 중심. 영구형 주거와 구분 관리.'),
 'USM025': dict(hq=H('Phoenix','AZ',33.50,-112.00), roles=['lessor','provider'], primary='lessor', types=['Volumetric'], regions=['Nationwide'],
   fin=dict(periodEnd='2025-12-31', fy='FY2025 (CY2025)', rev=2281.4, op=181.5, cfo=762.0, capex=273.2, conf='v',
            src='https://seekingalpha.com/pr/20406771',
            note='WillScot Holdings 연결(모듈러 공간+포터블 스토리지). Capex는 순 임대장비 투자. 조정 EBITDA $971.0M.'),
   ov='북미 최대 이동식 모듈러 공간·포터블 스토리지 임대 업체. 순수 영구 주거 모듈러 업체는 아니며 상업·교육·워크포스 중심.'),
 'USM026': dict(hq=H('Plymouth','MN',45.01,-93.46,'e'), roles=['lessor','provider'], primary='lessor', types=['Volumetric'], regions=['Midwest','Nationwide'], ov='상업·교육·의료·워크포스용 이동식 모듈러 임대. 별도 세그먼트로 관리.'),
 'USM027': dict(hq=H('Calgary (캐나다)','AB',None,None,'e'), roles=['lessor','provider','gc'], primary='lessor', types=['Volumetric','Hybrid'], regions=['Nationwide'], ov='캐나다 ATCO Ltd. 산하 워크포스·기관용 모듈러(Structures & Logistics). 미국 사업은 연결실적에서 분리 필요. 미국 소재지 확인 전까지 지도 미표시.'),
 'USM028': dict(hq=H('해외 본사','--',None,None,'l'), roles=['lessor','gc'], primary='lessor', types=['Volumetric','Hybrid'], regions=['Nationwide'], ov='국제 워크포스 숙소·EPC. 미국 활동 입증 시에만 포함(엑셀 지침). 지도 미표시.'),
 'USM029': dict(hq=H('Appleton','WI',44.26,-88.41), roles=['gc','provider'], primary='gc', types=['Volumetric','Panelized','Prefabricated'], regions=['Midwest'],
   fac=dict(count='1+', conf='e', summary='Bildt 모듈러·프리팹 제조사(Appleton, WI, 2025-01 발표). 현재 가동 상태 검증 필요.', list=[('Appleton, WI (Bildt)',44.26,-88.41)]),
   ov='위스콘신 기반 종합건설사 Boldt가 2025년 1월 설립한 모듈러·프리팹 제조 자회사 Bildt를 운영. 의료·기관·멀티패밀리 확대 가능성.'),
 'USM030': dict(hq=H('Minneapolis','MN',44.98,-93.27), roles=['gc','developer'], primary='gc', types=['Hybrid','Panelized','Volumetric (project)'], regions=['Nationwide','Midwest'], ov='전국 종합건설·디자인빌드사. 멀티패밀리·의료·교육·데이터센터 등에서 프리팹/오프사이트를 프로젝트별 적용. 확인된 모듈러 범위만 추적.'),
 'USM031': dict(hq=H('Omaha','NE',41.26,-95.93), roles=['gc'], primary='gc', types=['Hybrid','Prefabricated components'], regions=['Nationwide'], ov='인프라·산업·기관 중심 전국 종합건설/EPC. 순수 주거 모듈러 업체는 아님.'),
 'USM032': dict(hq=H('New York','NY',40.75,-73.99), roles=['gc'], primary='gc', types=['Hybrid','Panelized'], regions=['Nationwide','Northeast'], ov='전국 종합건설사(Hochtief/ACS 계열). 모듈러가 명시된 프로젝트만 추적.'),
 'USM033': dict(hq=H('New York','NY',40.74,-74.02), roles=['gc','developer'], primary='gc', types=['Hybrid','Panelized'], regions=['Northeast','Southeast','West','Nationwide'], ov='Skanska 그룹 미국 법인. 일부 시장에서 개발사업 병행. 그룹 재무를 모듈러에 직접 귀속하지 않음.'),
 'USM034': dict(hq=H('New York','NY',40.76,-73.98), roles=['developer','gc'], primary='developer', types=['Hybrid','Modular (project)'], regions=['Northeast','West','Nationwide'], ov='Lendlease 그룹의 미주 법인. 대도시 개발·시공·투자관리. 모듈러 명시 프로젝트만 추적.'),
 'USM035': dict(hq=H('Charleston','SC',32.78,-79.93), roles=['developer'], primary='developer', types=['Modular/offsite (project)'], regions=['Sun Belt','Nationwide'], ov='세계 최대급 임대주택 운용·개발사. 모듈러 제조사가 아닌 수요측(발주) 개발사로 분류하며 프로젝트 증거가 있을 때만 반영.'),
 'USM036': dict(hq=H('Miami','FL',25.80,-80.20),
   roles=['builder','developer'], primary='builder', types=['Site-built (offsite by project)'], regions=['Sun Belt','Nationwide'],
   fin=dict(opLabel='영업이익(산출)', fy='FY2025 (~2025-11-30)', rev=34186.9, op=2896.4, cfo=None, capex=None, conf='v',
            src='https://www.sec.gov/Archives/edgar/data/920760/000162828025057408/ex991-20251130x8kq4.htm',
            note='영업이익은 세그먼트 합계(주택 3,015.3+금융 612.5+멀티패밀리 −75.5+기타 −19.1)에서 본사 G&A 636.7을 차감한 산출치(재단 기부금 제외). 직접 공시 합계 아님.'),
   ov='미국 최대급 상장 홈빌더. 현장시공 중심이나 프로젝트별 오프사이트 적용. 모듈러 업체의 잠재 고객·파트너.'),
 'USM037': dict(hq=H('Arlington','TX',32.74,-97.11), roles=['builder','developer'], primary='builder', types=['Site-built (offsite by project)'], regions=['Sun Belt','Nationwide'],
   fin=dict(periodEnd='2025-09-30', fy='FY2025 (~2025-09-30)', rev=34250.4, op=4739.9, cfo=3420.9, capex=None, conf='v', opLabel='세전이익(대용)',
            src='https://www.sec.gov/Archives/edgar/data/882184/000088218425000044/a9302025exhibit991.htm', note='영업이익 대신 세전이익 사용.'),
   ov='판매량 기준 미국 최대 홈빌더. 토지개발·BTR 포함. 모듈러 사용은 프로젝트 검증 시에만 반영.'),
 'USM038': dict(hq=H('Atlanta','GA',33.75,-84.39), roles=['builder','developer'], primary='builder', types=['Site-built (offsite by project)'], regions=['Sun Belt','Nationwide'],
   fin=dict(periodEnd='2025-12-31', fy='FY2025 (CY2025)', rev=17312.0, op=2911.3, cfo=1871.2, capex=None, conf='v', opLabel='세전이익(대용)',
            src='https://s204.q4cdn.com/680895981/files/doc_earnings/2025/q4/earnings-result/Q4-2025-Earnings-Release.pdf', note='영업이익 대신 세전이익 사용.'),
   ov='전국 단독·액티브 어덜트 주택 빌더. 인접 시장 벤치마크/잠재 고객.'),
 'USM039': dict(hq=H('Reston','VA',38.96,-77.34), roles=['builder'], primary='builder', types=['Site-built (component sourcing)'], regions=['Mid-Atlantic','Northeast','Midwest'],
   fin=dict(periodEnd='2025-12-31', fy='FY2025 (CY2025)', rev=10324.0, op=1761.9, cfo=None, capex=None, conf='v', opLabel='세전이익(대용)',
            src='https://www.sec.gov/Archives/edgar/data/906163/000090616326000004/q42025ex991_earningsrelease.htm', note='영업이익 대신 세전이익 사용. 토지 경량(land-light) 모델.'),
   ov='동부·중부대서양 중심 토지경량 홈빌더. 공급망·고객 벤치마크 성격.'),
 'USM040': dict(hq=H('Los Angeles','CA',34.05,-118.24), roles=['builder','developer'], primary='builder', types=['Site-built (offsite by project)'], regions=['West','Southwest','Sun Belt'],
   fin=dict(fy='FY2025 (~2025-11-30)', rev=6236.2, op=507.1, cfo=None, capex=None, conf='v', opLabel='주택사업 영업이익',
            src='https://www.sec.gov/Archives/edgar/data/795266/000079526625000109/exh991kbh-earningsrelease1.htm', note='Homebuilding operating income 기준.'),
   ov='서부·남서부·선벨트 주문형 주택 빌더. 확인된 모듈러 조달만 추적.'),
 'USM041': dict(hq=H('Oakland','CA',37.80,-122.27,'e'), roles=['tech','panel'], primary='tech', types=['Panelized','3D-printed components'], regions=['West'], ov='3D 프린팅 패널 기반 ADU·주택 스타트업. 사업 지속 여부(운영 중단 보도 가능성)를 우선 검증해야 함.'),
 'USM042': dict(hq=H('Los Angeles','CA',34.02,-118.35,'e'), roles=['tech','provider'], primary='tech', types=['3D-printed components','Hybrid'], regions=['West'], ov='재활용 폴리머 3D프린팅 모듈 기반 백야드 스튜디오·ADU. 전통 볼류메트릭과 구분되는 인접 기술.'),
}

# 지식기반 직원 추정 (conf='e')
EMP_EST = {
 'USM024': '약 1,000+ (모회사 기준 추정)', 'USM025': '약 4,500 (추정)',
 'USM031': '약 3만 (추정)', 'USM035': '약 2.5만 (추정)',
 'USM036': '약 1.3만 (추정)', 'USM037': '약 1.4만 (추정)', 'USM038': '약 6,000+ (추정)', 'USM039': '약 6,000+ (추정)',
}

def load():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    def rows(name):
        ws = wb[name]; it = ws.iter_rows(values_only=True); h = next(it)
        return [dict(zip(h, r)) for r in it if any(c is not None for c in r)]
    return rows('CompanyMaster'), rows('Financials'), rows('Factories'), rows('Projects'), rows('Sources'), rows('MarketBenchmarks')

def split(s): return [x.strip() for x in re.split(r'[;/]\s*', s or '') if x.strip()] if s else []

def main():
    cm, fin, fac, prj, src, mb = load()
    out = []
    for c in cm:
        cid = c['CompanyID']; x = X.get(cid, {})
        # 엑셀 컬럼 오프셋 보정: 일부 행은 컬럼이 한 칸 당겨져 있음(예: USM004~)
        shifted = c.get('ResearchNotes') is None and c.get('InitialConfidence') is not None and str(c.get('InitialConfidence','')).startswith('http')
        if shifted:
            reg, endm, fds, conf, url, note = c['ProductRegulatoryType'], c['Regions'], c['EndMarkets'], c['FinancialDisclosureStatus'], c['InitialConfidence'], c['OfficialOrPrimarySourceURL']
            c = dict(c, ProductRegulatoryType=None, Regions=reg, EndMarkets=endm, FinancialDisclosureStatus=fds, InitialConfidence=conf, OfficialOrPrimarySourceURL=url, ResearchNotes=note)
        f = x.get('fin')
        if f:
            op, rev = f.get('op'), f.get('rev')
            f['margin'] = round(op / rev * 100, 1) if op is not None and rev else None
            f.setdefault('opLabel', '영업이익')
        emp = x.get('emp')
        if emp is None and cid in EMP_EST: emp = dict(v=EMP_EST[cid], conf='e', note='지식 기반 추정, 공시 확인 필요')
        if emp is None: emp = dict(v=None, conf=None, note='미확인')
        facd = x.get('fac') or dict(count=None, conf=None, summary='공장 정보 미확인(업데이트 대기). 엑셀 비고: ' + (c.get('ResearchNotes') or '-'), list=[])
        facd['list'] = [dict(name=n, lat=la, lon=lo) for (n, la, lo) in facd.get('list', [])]
        # 엑셀 Factories 반영
        xl_fac = [dict(name=r['FacilityName'], loc=r['CityOrCountry'], count=r['FacilityCount'], note=r['DataQualityNote'], src=r['SourceURL']) for r in fac if r['CompanyID'] == cid]
        projs = [dict(name=p['ProjectName'], city=p['City'], st=p['State'] or '', market=p['EndMarket'], method=p['ConstructionMethod'], scale=p['Scale'], stage=p['ProjectStage'], note=p['ScopeAndCaveats'], src=p['SourceURL']) for p in prj if p['CompanyID'] == cid]
        projs += x.get('projects', [])
        out.append(dict(
            id=cid, name=c['CompanyName'], legal=c['LegalOrParentName'], ownership=c['OwnershipType'],
            rolesRaw=c['BusinessRoles'], roles=x.get('roles', []), primary=x.get('primary', 'provider'),
            methodRaw=c['ManufacturingMethod'], types=x.get('types', split(c['ManufacturingMethod'])),
            regionsRaw=c['Regions'], regions=x.get('regions', []), endMarkets=split(c['EndMarkets']),
            hq=x.get('hq'), fin=f, finStatus=c['FinancialDisclosureStatus'], emp=emp, fac=facd, xlFac=xl_fac,
            overview=x.get('ov', ''), portfolio=x.get('port', split(c['EndMarkets'])), brands=x.get('brands', []),
            projects=projs, confidence=c['InitialConfidence'], url=c['OfficialOrPrimarySourceURL'], notes=c['ResearchNotes'], extra=x.get('note'), finNote=x.get('fin_note'),
        ))
    data = dict(companies=out, roles=ROLE, sources=src, market=mb, generatedAt=None)
    return data

if __name__ == '__main__':
    d = main()
    json.dump(d, open(sys.argv[2], 'w'), ensure_ascii=False, indent=1)
    print(len(d['companies']))

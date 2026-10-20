# US Modular Housing Intelligence

## 구성
- `index.html` : 단일 파일 대시보드(데이터·지도 내장, 로그인 없음)
- `data/data.json` : 데이터 원본 (웹으로 열면 화면이 매일 00:00 KST 에 이 파일을 다시 읽음)
- `update_data.py` : [1단계] SEC EDGAR XBRL 로 상장사/SEC 보고기업 재무 갱신
- `collect.py` : [2단계] SEC 10-K 본문, 공식 사이트(robots.txt 준수), GDELT/Google News RSS, Wikidata 수집·병합
- `tests/` : 오프라인 단위 테스트(샘플 HTML/JSON 사용)
- `.github/workflows/daily-update.yml` : 매일 23:30 KST(=14:30 UTC) 실행 → 00:00 에 새 데이터 준비
- `build/` : 엑셀 → JSON 변환, 지도 경로, 재빌드

## 설정(필수, GitHub 저장소 Settings > Secrets and variables > Actions)
- `SEC_UA` : "이름 이메일"  (SEC 요구사항)
- `CRAWL_UA` : "ModularDashboardBot/1.0 (contact: 이메일)"  (사이트 크롤링용)

## 원칙
- 유료 API·AI 호출 없음. 모든 자동값에 출처 URL·수집일·`auto` 표시.
- 수동 입력 공시확인값은 더 새로운 공시가 없으면 덮어쓰지 않음. 이상값(기존 대비 0.5~2배 밖)은 보류하고 `data/auto_log.json` 에 기록.
- 자동수집 프로젝트는 '자동수집(미검증)' 으로만 추가, 삭제하지 않음.
- 비상장사 매출·이익·현금흐름은 공개 출처가 없어 자동수집하지 않음(공란 = 미확인).

## 수동 실행
`SEC_UA="이름 이메일" python3 collect.py [--only USM008 USM010]`

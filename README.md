# US Modular Housing Intelligence
- `index.html` : 단일 파일 대시보드(데이터·지도 내장, 더블클릭으로 열림, 로그인 없음)
- `data/data.json` : 데이터 원본. 웹서버/GitHub Pages 등에 올리면 페이지가 매일 00:01 KST에 이 파일을 다시 읽어 화면을 갱신
- `update_data.py` + `.github/workflows/daily-update.yml` : 매일 00:01 KST에 SEC EDGAR 재무를 갱신하고 커밋(배포)
- `build/` : 엑셀→JSON 변환(data_src.py), 지도 경로, 재빌드(build.py)

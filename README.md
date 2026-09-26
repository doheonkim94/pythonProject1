# PASTA 마케팅 자동화 — Claude Code 루틴 전환 가이드

## 왜 Claude Code로 전환하는가
지금까지 채팅에서는 구글 드라이브 MCP 커넥터로 시트를 "파일 전체 텍스트 변환" 방식으로 읽어서,
큰 시트(10개 탭, 그중 하나만 25만 자)의 뒷부분(9월 데이터)에 도달하지 못하고 잘렸습니다.

Claude Code에서는 실제 파이썬 코드로 구글시트 API(`spreadsheets().values().get(range=...)`)를
직접 호출합니다. 이건 파일 전체가 아니라 **필요한 범위만** 요청하는 방식이라, 시트가 아무리 커도
오늘/전일/전주 동요일 세 행만 정확히 가져옵니다. 지금까지 막혔던 문제의 근본 해결책입니다.

## v2 변경사항 (중요)
v1 스크립트는 노션 데이터베이스의 "속성(숫자 컬럼)"만 채우고, 채팅에서 보여드린 6단 구성 본문
(매체별 효율 표, 소재 하이라이트 표+드라이브 링크, 개선점)은 만들지 않았습니다. 그래서 실제로
돌리면 속성은 찼는데 본문은 비어있는, 예시와 다른 리포트가 나왔을 겁니다. v2는 아래를 추가해서
본문까지 동일하게 만듭니다:
- `MediaMix_YY.MM.` 탭에서 매체별 예산/비중/예상 CPI·CPA를 읽어 3번 섹션에 넣음
- `신규 소재 성과` 탭에서 해당 날짜에 운영 중이던(D+14 기간이 그 날짜를 포함하는) 소재를 골라
  고효율/저효율로 나누고, 4번 섹션에 표로 넣음
- 그 소재들의 실제 구글 드라이브 원본 링크를 찾아 소재명에 연결
- 위 데이터를 근거로 6번 개선점을 자동 생성 (저효율 소재 목록, 고효율 패턴 등)

이 때문에 **구글 서비스 계정에 Drive API도 추가로 켜야 하고**, 소재 원본 파일이 있는
드라이브 폴더도 공유해야 합니다 (아래 1번 단계에 포함).

## 1. 구글 서비스 계정 만들기 (한 번만)
1. https://console.cloud.google.com 에서 프로젝트 선택(또는 생성)
2. "API 및 서비스 > 라이브러리"에서 **Google Sheets API** 와 **Google Drive API** 둘 다 사용 설정
   (v2부터는 소재 원본 링크를 찾기 위해 Drive API도 필요합니다)
3. "API 및 서비스 > 사용자 인증 정보 > 사용자 인증 정보 만들기 > 서비스 계정" 으로 서비스 계정 생성
4. 생성된 서비스 계정 > "키" 탭 > "키 추가 > 새 키 만들기" > JSON 선택 → 파일 다운로드
5. 다운로드한 JSON을 열어 `client_email` 값을 복사 (예: `xxx@xxx.iam.gserviceaccount.com`)
6. 대상 구글시트("[🟡리포트]_파스타_(KHC-SMC) 2026.")를 열어 **공유** > 위 이메일을 **뷰어**로 추가
7. 소재 원본 파일이 있는 구글 드라이브 폴더도 같은 이메일을 **뷰어**로 공유
   (신규 소재 성과 탭의 소재명과 파일명이 정확히 일치하는 그 폴더)

## 2. Notion 통합 토큰 만들기 (한 번만)
1. https://www.notion.so/my-integrations 에서 "New integration" 생성
2. 발급된 토큰(`ntn_...`로 시작) 복사
3. 노션에서 "데일리 퍼포먼스 리포트" 페이지 열기 > 우측 상단 `...` > Connections > 방금 만든 통합 연결
   (하위의 일간/주간/월간 데이터베이스는 부모 페이지 연결을 상속합니다)

## 3. 구글챗 웹훅 만들기 (한 번만)
알림 받을 스페이스 > 스페이스 설정 > 앱 및 통합 > 웹훅 추가 → URL 복사

## 4. Claude Code 루틴 등록
1. 이 저장소를 GitHub에 올린다
2. https://claude.ai/code/routines 접속 > New routine
3. Repository: 방금 만든 저장소
4. Environment: 아래 세 값을 **API credentials**(환경변수 아님, 값이 노출되지 않는 자격증명 섹션)로 등록
   - `GOOGLE_SERVICE_ACCOUNT_JSON` — 1번에서 받은 JSON 파일의 전체 내용을 그대로 붙여넣기
   - `NOTION_TOKEN` — 2번 토큰
   - `GCHAT_WEBHOOK_URL` — 3번 URL
5. Prompt 예시:
   > `daily_report.py`를 실행해서 오늘자 파스타 일간 리포트를 만들어줘. 에러가 나면 원인과 함께 실패로 보고해.
6. Trigger: Scheduled > Daily, 원하는 시간(예: 매일 오전 8시)
7. 저장 전에 **Run now**로 먼저 한 번 수동 실행해서 결과 확인 (특히 `--date`를 과거 날짜로 지정해서
   이미 대행사가 채워둔 날짜로 먼저 테스트하는 걸 추천 — `python daily_report.py --date 2026-09-20`)

## 로컬에서 먼저 테스트하고 싶다면
```bash
pip install -r requirements.txt
export GOOGLE_SERVICE_ACCOUNT_JSON="$(cat /path/to/service-account.json)"
export NOTION_TOKEN="ntn_..."
export GCHAT_WEBHOOK_URL="https://chat.googleapis.com/..."
python daily_report.py --date 2026-09-20
```

## 다음 단계 (확장)
- 주간/월간 리포트: 같은 패턴으로 `weekly_report.py`, `monthly_report.py`를 추가하고,
  각각 주 1회(월요일)/월 1회(1일) 스케줄 트리거를 가진 별도 루틴으로 등록
- 매체별 실적: Windsor.ai 연동 후 `fetch_rows` 대신 Windsor.ai API 호출로 대체하면
  대행사 시트 의존 없이 매체에서 직접 가져올 수 있음
- 소재 하이라이트: `신규 소재 성과` 탭도 같은 방식(`values().get`)으로 범위를 지정해 읽어와
  리포트 본문에 추가 가능

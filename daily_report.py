#!/usr/bin/env python3
"""
PASTA 일간 퍼포먼스 리포트 자동화 스크립트
================================================
Claude Code 루틴(routine)이 매일 아침 이 스크립트를 실행하는 것을 전제로 작성했습니다.

이 스크립트가 하는 일:
  1. 구글시트 "[🟡리포트]_파스타_(KHC-SMC) 2026." 의 '일별 Summary - Total' 탭에서
     ① 오늘 행 ② 전일 행 ③ 전주 동요일 행 딱 3개 행만 정확히 골라 읽는다.
     (Drive MCP 커넥터의 "파일 전체를 텍스트로 변환" 방식과 달리, 구글시트 API의
     values().get(range=...) 로 필요한 범위만 요청하므로 파일 크기와 무관하게 동작한다.)
  2. 전일 대비·전주 동요일 대비 증감률을 계산한다.
  3. Notion '일간 리포트' 데이터베이스에 오늘자 행을 생성한다.
  4. 생성된 노션 페이지 URL을 구글챗 웹훅으로 알림 전송한다.

필요한 환경변수 (Claude Code 루틴의 "환경 변수" 또는 "API credentials"에 등록):
  GOOGLE_SERVICE_ACCOUNT_JSON   구글 서비스 계정 키 파일의 전체 JSON 내용 (문자열)
  NOTION_TOKEN                  Notion 내부 통합(integration) 토큰
  GCHAT_WEBHOOK_URL              구글챗 수신 웹훅 URL

사전 준비 (한 번만):
  1) Google Cloud Console에서 서비스 계정을 만들고 Google Sheets API를 켠다.
  2) 서비스 계정 키(JSON)를 발급받는다.
  3) 그 키 안의 client_email(예: xxx@xxx.iam.gserviceaccount.com)을
     대상 구글시트에 "뷰어"로 공유한다. ← 이 한 줄이 이 스크립트가 동작하기 위한 전제.
  4) notion.so/my-integrations 에서 내부 통합을 만들고 토큰을 받은 뒤,
     '데일리 퍼포먼스 리포트' 페이지에 그 통합을 연결(Connect)한다.
  5) 구글챗 스페이스에서 웹훅 URL을 만든다 (스페이스 설정 > 앱 및 통합 > 웹훅).

사용법:
  python daily_report.py                  # 오늘 날짜로 실행
  python daily_report.py --date 2026-09-01  # 특정 날짜로 실행 (백필/테스트용)
"""

import argparse
import datetime
import json
import os
import sys

import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── 고정 설정: 이 프로젝트의 구글시트/노션 리소스 ID ─────────────────────────
SHEET_ID = "1iCeRn5-bPPFR47RydluurDcCBrKlYlMxwW02q5Bb9rA"
SHEET_TAB = "일별 Summary - Total"

NOTION_API = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"
DAILY_DB_ID = "17ff48f0-3fd9-4ab0-9c7c-5d64bda5aa5d"  # 일간 리포트 데이터 소스 ID

WEEKDAY_KR = ["월", "화", "수", "목", "금", "토", "일"]

# 시트 원본 컬럼명 → 노션 프로퍼티명 매핑. 시트 헤더 문구가 바뀌면 여기만 고치면 된다.
COLUMN_MAP = {
    "집행 금액": "집행 금액(원)",
    "예산 소진율": "예산 소진율(%)",
    "앱설치 (Total)": "앱설치",
    "회원가입 (Total)": "회원가입",
    "설치당 단가 (Total)": "CPI(원)",
    "회원가입당 단가 (Total)": "CPA(원)",
    "가입 전환율 (Total)": "가입전환율(%)",
}


def sheets_client():
    """서비스 계정 인증으로 Sheets API v4 클라이언트를 만든다."""
    raw = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    info = json.loads(raw)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    return build("sheets", "v4", credentials=creds)


def find_header_row(values):
    """'일' 로 시작하고 '예산' 을 포함하는 행을 헤더 행으로 판단한다.
    이 시트는 상단에 제목/안내문이 여러 줄 있어서 고정 행 번호를 쓰면 깨지기 쉽다."""
    for i, row in enumerate(values):
        if row and row[0].strip() == "일" and any("예산" in c for c in row):
            return i
    raise RuntimeError("헤더 행을 못 찾았습니다 — 시트 상단 구조가 바뀐 것 같습니다.")


def parse_number(cell: str):
    """'  1,234 ' / '-' / '12.3%' 같은 셀 값을 숫자로 정리한다."""
    if cell is None:
        return None
    s = cell.strip().replace(",", "")
    if s in ("", "-"):
        return None
    if s.endswith("%"):
        try:
            return round(float(s[:-1]) / 100, 4)
        except ValueError:
            return None
    try:
        return float(s) if "." in s else int(s)
    except ValueError:
        return None


def row_date_label(d: datetime.date) -> str:
    """이 시트가 쓰는 날짜 표기('9/1 (화)')로 변환. 월/일 앞에 0을 붙이지 않는다."""
    return f"{d.month}/{d.day} ({WEEKDAY_KR[d.weekday()]})"


def fetch_rows(service, target_date: datetime.date):
    """오늘/전일/전주 동요일 세 날짜의 데이터 행을 한 번의 API 호출로 가져온다."""
    resp = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=SHEET_ID, range=f"'{SHEET_TAB}'!A1:BN2000")
        .execute()
    )
    values = resp.get("values", [])
    header_idx = find_header_row(values)
    header = values[header_idx]

    wanted = {
        "today": row_date_label(target_date),
        "prev_day": row_date_label(target_date - datetime.timedelta(days=1)),
        "prev_week_same_weekday": row_date_label(target_date - datetime.timedelta(days=7)),
    }

    found = {}
    for row in values[header_idx + 1 :]:
        if not row:
            continue
        label = row[0].strip()
        for key, date_label in wanted.items():
            if label == date_label and key not in found:
                # 셀 개수가 헤더보다 짧을 수 있으니 채워서 zip
                padded = row + [""] * (len(header) - len(row))
                found[key] = dict(zip(header, padded))

    missing = [k for k in wanted if k not in found]
    if missing:
        raise RuntimeError(
            f"다음 날짜 행을 시트에서 못 찾았습니다: {[wanted[k] for k in missing]} "
            "— 아직 대행사가 그 날짜분을 안 채웠을 수 있습니다."
        )
    return found


def pct_change(new, old):
    if new is None or old in (None, 0):
        return None
    return round((new - old) / old, 4)


def build_notion_properties(target_date, today, prev_day, prev_week):
    def g(row, sheet_col):
        return parse_number(row.get(sheet_col))

    설치_today = g(today, "앱설치 (Total)")
    가입_today = g(today, "회원가입 (Total)")
    설치_prev = g(prev_day, "앱설치 (Total)")
    가입_prev = g(prev_day, "회원가입 (Total)")
    설치_pw = g(prev_week, "앱설치 (Total)")
    가입_pw = g(prev_week, "회원가입 (Total)")

    props = {
        "날짜": {"title": [{"text": {"content": target_date.isoformat()}}]},
        "요일": {"select": {"name": WEEKDAY_KR[target_date.weekday()]}},
        "집행 금액(원)": {"number": g(today, "집행 금액")},
        "예산 소진율(%)": {"number": g(today, "예산 소진율")},
        "앱설치": {"number": 설치_today},
        "회원가입": {"number": 가입_today},
        "CPI(원)": {"number": g(today, "설치당 단가 (Total)")},
        "CPA(원)": {"number": g(today, "회원가입당 단가 (Total)")},
        "가입전환율(%)": {"number": g(today, "가입 전환율 (Total)")},
        "전일 대비 설치 증감(%)": {"number": pct_change(설치_today, 설치_prev)},
        "전일 대비 가입 증감(%)": {"number": pct_change(가입_today, 가입_prev)},
        "전주 동요일 대비 설치 증감(%)": {"number": pct_change(설치_today, 설치_pw)},
        "전주 동요일 대비 가입 증감(%)": {"number": pct_change(가입_today, 가입_pw)},
    }
    return props


def create_notion_page(properties):
    headers = {
        "Authorization": f"Bearer {os.environ['NOTION_TOKEN']}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }
    payload = {
        "parent": {"data_source_id": DAILY_DB_ID},
        "properties": properties,
    }
    resp = requests.post(f"{NOTION_API}/pages", headers=headers, json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json()


def notify_gchat(page_url: str, target_date: datetime.date):
    webhook = os.environ.get("GCHAT_WEBHOOK_URL")
    if not webhook:
        print("GCHAT_WEBHOOK_URL 미설정 — 구글챗 알림은 건너뜁니다.")
        return
    text = f"📊 {target_date.isoformat()} 파스타 일간 리포트가 생성됐습니다.\n{page_url}"
    resp = requests.post(webhook, json={"text": text}, timeout=15)
    resp.raise_for_status()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="YYYY-MM-DD, 기본값은 오늘(KST 기준은 호출부에서 맞출 것)")
    args = parser.parse_args()

    target_date = (
        datetime.date.fromisoformat(args.date) if args.date else datetime.date.today()
    )

    service = sheets_client()
    rows = fetch_rows(service, target_date)

    properties = build_notion_properties(
        target_date, rows["today"], rows["prev_day"], rows["prev_week_same_weekday"]
    )
    page = create_notion_page(properties)
    page_url = page.get("url", "")
    print(f"노션 페이지 생성 완료: {page_url}")

    notify_gchat(page_url, target_date)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001 — 루틴 로그에 그대로 남기기 위해 넓게 잡음
        print(f"실패: {exc}", file=sys.stderr)
        sys.exit(1)

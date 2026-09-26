#!/usr/bin/env python3
"""
경쟁사(다이어트 앱) 광고소재/콘텐츠 모니터링 스크립트
====================================================
대상 스프레드시트: [파스타] 경쟁사_광고소재_모니터링
  https://docs.google.com/spreadsheets/d/1_jkR8Nh9WwWS7selKoC6Am6wHhIhmMs1fQgTRsuGmoU
대상 브랜드: 필라이즈, 닥터다이어리, 인아웃

수집 소스:
  1) 네이버 검색 API(블로그/카페/뉴스) — 이미 작동 확인됨
  2) Meta 광고 라이브러리 API — Meta 개발자 계정의 Ad Library API 접근 승인
     (facebook.com/ads/library/api 의 본인 인증 절차)이 끝나야 실제로 값이 들어온다.
     META_ACCESS_TOKEN이 없거나 그 승인이 안 끝난 상태면 이 부분만 건너뛴다.

두 소스 모두 "이미 시트에 있는 링크/광고ID는 다시 안 넣는다"는 방식으로 누적한다
(대행사 시트의 최초수집일 컬럼과 같은 개념).

필요한 환경변수:
  GOOGLE_SERVICE_ACCOUNT_JSON   구글 서비스 계정 키 JSON (문자열)
  NAVER_CLIENT_ID               네이버 검색 API 클라이언트 ID
  NAVER_CLIENT_SECRET           네이버 검색 API 클라이언트 시크릿
  META_ACCESS_TOKEN             Meta Ad Library API 액세스 토큰 (선택 — 없으면 그 부분만 생략)

사용법:
  python competitor_report.py
"""

import datetime
import json
import os
import sys
import time

import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

SPREADSHEET_ID = "1_jkR8Nh9WwWS7selKoC6Am6wHhIhmMs1fQgTRsuGmoU"

# 브랜드 → 검색에 쓸 키워드 목록 (브랜드명과 검색어가 다를 때를 대비해 리스트로 둔다)
BRANDS = {
    "필라이즈": ["필라이즈"],
    "닥터다이어리": ["닥터다이어리"],
    "인아웃": ["인아웃"],
}

NAVER_TABS = {
    "blog": ("네이버블로그", ["게재일", "채널", "브랜드", "키워드", "순위", "제목", "심의", "심의번호", "요약", "링크", "최초수집일"]),
    "cafearticle": ("네이버카페", ["게재일", "채널", "브랜드", "키워드", "순위", "제목", "출처", "요약", "링크", "최초수집일"]),
    "news": ("네이버뉴스", ["게재일", "채널", "브랜드", "키워드", "순위", "제목", "출처", "요약", "링크", "최초수집일"]),
}
CHANNEL_LABEL = {"blog": "네이버블로그", "cafearticle": "네이버카페", "news": "네이버뉴스"}

META_TAB_HEADER = ["최초수집일", "상태", "출처", "광고ID", "광고주", "소재", "유형", "문구", "광고 시작일", "랜딩 URL", "광고 링크"]


def sheets_client():
    info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds)


def strip_html(s):
    return (s or "").replace("<b>", "").replace("</b>", "")


def read_column(service, tab, column_name):
    """해당 탭의 헤더에서 column_name 열의 기존 값 전체를 집합으로 읽어온다 (중복 판단용)."""
    resp = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{tab}'!A1:Z5000"
    ).execute()
    values = resp.get("values", [])
    if not values:
        return set()
    header = values[0]
    if column_name not in header:
        return set()
    idx = header.index(column_name)
    return {row[idx] for row in values[1:] if idx < len(row) and row[idx]}


def append_rows(service, tab, rows):
    if not rows:
        return
    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{tab}'!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()


# ── 1. 네이버 블로그/카페/뉴스 ───────────────────────────────────────────
def collect_naver(service, today):
    client_id = os.environ["NAVER_CLIENT_ID"]
    client_secret = os.environ["NAVER_CLIENT_SECRET"]
    headers = {"X-Naver-Client-Id": client_id, "X-Naver-Client-Secret": client_secret}

    for kind, (tab, _header) in NAVER_TABS.items():
        seen_links = read_column(service, tab, "링크")
        new_rows = []
        for brand, keywords in BRANDS.items():
            for kw in keywords:
                resp = requests.get(
                    f"https://openapi.naver.com/v1/search/{kind}.json",
                    headers=headers,
                    params={"query": kw, "display": 20, "sort": "date"},
                    timeout=10,
                )
                resp.raise_for_status()
                for rank, item in enumerate(resp.json().get("items", []), start=1):
                    link = item.get("link") or ""
                    if not link or link in seen_links:
                        continue
                    seen_links.add(link)

                    title = strip_html(item.get("title"))
                    desc = strip_html(item.get("description"))
                    pub = item.get("postdate") or item.get("pubDate", "")
                    if kind == "news" and pub:
                        try:
                            pub = datetime.datetime.strptime(pub, "%a, %d %b %Y %H:%M:%S %z").strftime("%Y-%m-%d")
                        except ValueError:
                            pass
                    elif kind == "blog" and pub and len(pub) == 8:
                        pub = f"{pub[:4]}-{pub[4:6]}-{pub[6:]}"

                    if kind == "blog":
                        row = [pub, tab, brand, kw, rank, title, "X", "", desc, link, today]
                    elif kind == "cafearticle":
                        row = [pub, tab, brand, kw, rank, title, "", desc, link, today]
                    else:  # news
                        orig = item.get("originallink", "")
                        source = orig.split("/")[2] if orig.count("/") >= 2 else ""
                        row = [pub, tab, brand, kw, rank, title, source, desc, link, today]
                    new_rows.append(row)
                time.sleep(0.2)  # 브랜드/키워드마다 호출 간 살짝 여유
        append_rows(service, tab, new_rows)
        print(f"[{tab}] 신규 {len(new_rows)}건 추가")


# ── 2. Meta 광고 라이브러리 ──────────────────────────────────────────────
def collect_meta(service, today):
    token = os.environ.get("META_ACCESS_TOKEN")
    if not token:
        print("[경고] META_ACCESS_TOKEN 미설정 — Meta 광고 라이브러리 수집은 건너뜁니다.")
        return

    for brand in BRANDS:
        tab = brand
        seen_ids = read_column(service, tab, "광고ID")
        try:
            resp = requests.get(
                "https://graph.facebook.com/v21.0/ads_archive",
                params={
                    "search_terms": brand,
                    "ad_reached_countries": '["KR"]',
                    "ad_active_status": "ALL",
                    "fields": "id,ad_creative_bodies,ad_creative_link_titles,ad_delivery_start_time,"
                              "ad_delivery_stop_time,ad_snapshot_url,page_name",
                    "access_token": token,
                    "limit": 50,
                },
                timeout=20,
            )
            resp.raise_for_status()
        except requests.HTTPError as exc:
            print(f"[경고] {brand} Meta Ad Library 조회 실패 ({exc}) — 건너뜁니다.")
            print(f"       ({resp.text[:300]})")
            continue

        new_rows = []
        for ad in resp.json().get("data", []):
            ad_id = ad.get("id")
            if not ad_id or ad_id in seen_ids:
                continue
            seen_ids.add(ad_id)

            status = "게재중" if not ad.get("ad_delivery_stop_time") else f"종료({ad['ad_delivery_stop_time'][:10]})"
            body = " / ".join(ad.get("ad_creative_bodies", []) or [])
            start = (ad.get("ad_delivery_start_time") or "")[:10]
            snapshot = ad.get("ad_snapshot_url", "")

            new_rows.append(
                [
                    today, status, "Meta", ad_id, ad.get("page_name", brand),
                    "", "", body, start, "", snapshot,
                ]
            )
        append_rows(service, tab, new_rows)
        print(f"[{tab}] Meta 신규 {len(new_rows)}건 추가")


def main():
    today = datetime.date.today().isoformat()
    service = sheets_client()
    collect_naver(service, today)
    collect_meta(service, today)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"실패: {exc}", file=sys.stderr)
        sys.exit(1)

#!/usr/bin/env python3
"""
소재 아이디어 생성용 인풋 요약 (로드맵 4단계)
============================================
아이디어 자체를 만들지는 않는다 — 이 스크립트는 아래 세 가지를 모아
읽기 쉬운 텍스트로 출력만 한다. 그 출력을 읽고 오늘의 소재 아이디어를
작성하는 건 이 스크립트를 실행하는 Claude Code 루틴(자동화 세션) 자신의
역할이다. (raw 데이터 → 서술형 판단은 스크립트가 아니라 Claude가 한다는
원래 설계 그대로.)

모으는 것:
  1. 소재 성과 DB(3단계)의 최근 고성과/저성과 소재 — 컨셉·인플루언서·CPA·CVR
  2. 경쟁사 모니터링 시트에서 오늘 새로 수집된 블로그/카페/뉴스 콘텐츠
  3. 오늘 날짜 기준 시즌/이벤트 힌트 (건강검진 시즌, 공휴일 등)

필요한 환경변수: GOOGLE_SERVICE_ACCOUNT_JSON, NOTION_TOKEN
사용법: python creative_idea_inputs.py
"""

import datetime
import os
import sys

import requests

import daily_report as dr

CREATIVE_DB_ID = "3e7fbc1a-b2c3-8145-bc0c-ff260fc43363"
COMPETITOR_SHEET_ID = "1_jkR8Nh9WwWS7selKoC6Am6wHhIhmMs1fQgTRsuGmoU"
COMPETITOR_NAVER_TABS = ["네이버블로그", "네이버카페", "네이버뉴스"]

# (시작 월,일) (끝 월,일) 힌트 — 대략적인 다이어트/헬스케어 업계 시즌성 참고용.
# 실제 캠페인 캘린더가 별도로 있다면 이 목록을 그걸로 교체하는 게 좋다.
SEASON_HINTS = [
    ((1, 1), (2, 28), "새해 다이어트 결심 시즌 성수기"),
    ((5, 1), (6, 30), "여름 대비 다이어트 시즌"),
    ((9, 1), (9, 30), "하반기 건강검진 시즌, 명절 전후 식습관 관리"),
    ((11, 1), (12, 31), "연말 다이어트/새해 목표 예열 시즌"),
]


def season_hint(d):
    for (sm, sd), (em, ed), hint in SEASON_HINTS:
        start, end = datetime.date(d.year, sm, sd), datetime.date(d.year, em, ed)
        if start <= d <= end:
            return hint
    return None


def notion_headers():
    return {
        "Authorization": f"Bearer {os.environ['NOTION_TOKEN']}",
        "Notion-Version": dr.NOTION_VERSION,
        "Content-Type": "application/json",
    }


def prop_text(page, name, kind="rich_text"):
    prop = page["properties"].get(name, {})
    if kind == "rich_text":
        return "".join(t["plain_text"] for t in prop.get("rich_text", []))
    if kind == "title":
        return "".join(t["plain_text"] for t in prop.get("title", []))
    if kind == "select":
        sel = prop.get("select")
        return sel["name"] if sel else ""
    if kind == "number":
        return prop.get("number")
    return None


def query_creative_db(tag, limit=5):
    payload = {
        "filter": {"property": "고성과여부", "select": {"equals": tag}},
        "sorts": [{"property": "최근갱신일", "direction": "descending"}],
        "page_size": limit,
    }
    resp = requests.post(
        f"{dr.NOTION_API}/databases/{CREATIVE_DB_ID}/query", headers=notion_headers(), json=payload, timeout=30
    )
    resp.raise_for_status()
    return resp.json().get("results", [])


def print_creative_summary():
    print("### 1. 소재 성과 DB — 최근 고성과 / 저성과")
    for tag in ["고성과", "저성과"]:
        pages = query_creative_db(tag)
        print(f"\n[{tag}] ({len(pages)}건)")
        for p in pages:
            name = prop_text(p, "소재명", "title")
            media = prop_text(p, "매체", "select")
            concept = prop_text(p, "컨셉")
            influencer = prop_text(p, "인플루언서")
            cpa = prop_text(p, "CPA", "number")
            cvr = prop_text(p, "가입전환율", "number")
            cvr_str = f"{cvr * 100:.1f}%" if cvr is not None else "-"
            print(f"- {name} [{media}] 컨셉:{concept or '-'} 인플루언서:{influencer or '-'} CPA:{cpa} CVR:{cvr_str}")


def print_competitor_summary():
    print("\n### 2. 경쟁사 최신 콘텐츠 (오늘 새로 수집된 것)")
    creds = dr.google_creds()
    sheets = dr.sheets_client(creds)
    today = datetime.date.today().isoformat()
    for tab in COMPETITOR_NAVER_TABS:
        resp = sheets.spreadsheets().values().get(
            spreadsheetId=COMPETITOR_SHEET_ID, range=f"'{tab}'!A1:K3000"
        ).execute()
        values = resp.get("values", [])
        if not values:
            continue
        header = values[0]
        try:
            idx_collect = header.index("최초수집일")
            idx_brand = header.index("브랜드")
            idx_title = header.index("제목")
        except ValueError:
            continue
        recent = [r for r in values[1:] if idx_collect < len(r) and r[idx_collect] == today]
        print(f"\n[{tab}] 오늘 신규 {len(recent)}건")
        for r in recent[:5]:
            brand = r[idx_brand] if idx_brand < len(r) else ""
            title = r[idx_title] if idx_title < len(r) else ""
            print(f"- ({brand}) {title}")


def print_season_hint():
    print("\n### 3. 시즌/캘린더 힌트")
    today = datetime.date.today()
    print(f"- 오늘: {today.isoformat()} ({dr.WEEKDAY_KR[today.weekday()]})")
    print(f"- 시즌 힌트: {season_hint(today) or '해당 없음'}")
    holiday_name = dr.KR_HOLIDAYS.get(today)
    if holiday_name:
        print(f"- 공휴일: {holiday_name}")


def main():
    print_creative_summary()
    print_competitor_summary()
    print_season_hint()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"실패: {exc}", file=sys.stderr)
        sys.exit(1)

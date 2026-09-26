#!/usr/bin/env python3
"""
소재 성과 DB 동기화 스크립트 (로드맵 3단계)
============================================
'신규 소재 성과' 탭에서 매체별(META/Tiktok/Moloco) 활성 소재를 읽어서,
노션 '소재 성과 DB'에 소재 단위로 누적·갱신한다. 매체 필터 순회, 고효율/
저효율 판정, 드라이브 링크 조회는 daily_report.py가 이미 가진 걸 그대로
가져다 쓴다 — 소재 데이터를 다시 파싱하는 코드를 새로 만들지 않는다.

동작:
  - 소재명이 DB에 없으면 새 페이지를 만들고 최초발견일을 오늘로 기록
  - 이미 있으면 최신 성과로 갱신하고 최근갱신일만 오늘로 갱신
  - 그날 매체별 rank_creatives() 결과(고효율/저효율)를 고성과여부 태그에 반영

이 DB는 4단계(소재 아이디어 자동 생성)의 핵심 인풋이 되도록 설계했다 —
소재별 컨셉(USP)·타깃(인플루언서)·성과지표가 한 곳에 시계열로 쌓인다.

필요한 환경변수 (daily_report.py와 동일):
  GOOGLE_SERVICE_ACCOUNT_JSON, NOTION_TOKEN

사용법:
  python creative_db_sync.py [--date YYYY-MM-DD]   # 기본값은 오늘
"""

import argparse
import datetime
import os
import sys

import requests

import daily_report as dr

CREATIVE_DB_ID = "3e7fbc1a-b2c3-8145-bc0c-ff260fc43363"


def notion_headers():
    return {
        "Authorization": f"Bearer {os.environ['NOTION_TOKEN']}",
        "Notion-Version": dr.NOTION_VERSION,
        "Content-Type": "application/json",
    }


def find_page_by_name(name, media):
    """같은 소재명이 여러 매체에 동시에 걸리는 경우가 있어서(같은 이미지를 META와
    Moloco에 함께 태우는 등), 소재명만으로는 중복 판정이 안 된다 — 매체까지 같아야
    같은 소재로 본다."""
    payload = {
        "filter": {
            "and": [
                {"property": "소재명", "title": {"equals": name}},
                {"property": "매체", "select": {"equals": media}},
            ]
        }
    }
    resp = requests.post(
        f"{dr.NOTION_API}/databases/{CREATIVE_DB_ID}/query",
        headers=notion_headers(), json=payload, timeout=30,
    )
    resp.raise_for_status()
    results = resp.json().get("results", [])
    return results[0]["id"] if results else None


def upsert_page(properties, page_id, today_str):
    if page_id:
        url = f"{dr.NOTION_API}/pages/{page_id}"
        resp = requests.patch(url, headers=notion_headers(), json={"properties": properties}, timeout=30)
    else:
        properties["최초발견일"] = {"date": {"start": today_str}}
        payload = {"parent": {"database_id": CREATIVE_DB_ID}, "properties": properties}
        resp = requests.post(f"{dr.NOTION_API}/pages", headers=notion_headers(), json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json()


def build_properties(media, creative, today_str, tag, link):
    cpi = round(creative["spend"] / creative["install"]) if creative.get("install") else None
    props = {
        "소재명": {"title": [{"text": {"content": creative["name"]}}]},
        "매체": {"select": {"name": media}},
        "컨셉": {"rich_text": [{"text": {"content": creative.get("usp") or ""}}]},
        "인플루언서": {"rich_text": [{"text": {"content": creative.get("influencer") or ""}}]},
        "운영기간": {"rich_text": [{"text": {"content": creative.get("period") or ""}}]},
        "집행금액": {"number": creative.get("spend")},
        "앱설치": {"number": creative.get("install")},
        "CPI": {"number": cpi},
        "회원가입": {"number": creative.get("signup")},
        "CPA": {"number": creative.get("cpa")},
        "가입전환율": {"number": creative.get("cvr")},
        "고성과여부": {"select": {"name": tag}},
        "최근갱신일": {"date": {"start": today_str}},
    }
    if link:
        props["드라이브 링크"] = {"url": link}
    return props


def sync(target_date):
    creds = dr.google_creds()
    sheets = dr.sheets_client(creds)
    drive = dr.drive_client(creds)
    today_str = datetime.date.today().isoformat()

    media_creatives = dr.fetch_creatives_by_media(sheets, target_date)

    created = updated = 0
    for media, creatives in media_creatives.items():
        high, low = dr.rank_creatives(creatives)
        high_names = {c["name"] for c in high}
        low_names = {c["name"] for c in low}

        for c in creatives:
            tag = "고성과" if c["name"] in high_names else "저성과" if c["name"] in low_names else "보통"
            try:
                link = dr.drive_link(drive, c["name"])
            except Exception as exc:  # noqa: BLE001
                print(f"[경고] {c['name']} 드라이브 조회 실패 ({exc})")
                link = None

            page_id = find_page_by_name(c["name"], media)
            props = build_properties(media, c, today_str, tag, link)
            upsert_page(props, page_id, today_str)
            if page_id:
                updated += 1
            else:
                created += 1
            print(f"[{media}] {c['name']} -> {tag} {'(갱신)' if page_id else '(신규)'}")

    print(f"완료: 신규 {created}건, 갱신 {updated}건")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="YYYY-MM-DD, 기본값은 오늘")
    args = parser.parse_args()
    target_date = datetime.date.fromisoformat(args.date) if args.date else datetime.date.today()
    sync(target_date)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"실패: {exc}", file=sys.stderr)
        sys.exit(1)

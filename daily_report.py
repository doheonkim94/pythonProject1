#!/usr/bin/env python3
"""
PASTA 일간 퍼포먼스 리포트 자동화 스크립트 (v2)
================================================
v1과의 차이: v1은 노션 데이터베이스의 "속성(properties)"만 채웠고, 실제 채팅에서 보여드린
6단 구성 본문(매체별 효율 표, 소재 하이라이트 표+링크, 개선점)은 만들지 않았습니다.
그래서 실제로 돌려보면 "속성은 찼는데 본문 내용이 비어있는" 다른 리포트가 나왔을 겁니다.
v2는 본문까지 동일한 6단 구성으로 생성합니다.

필요한 환경변수:
  GOOGLE_SERVICE_ACCOUNT_JSON   구글 서비스 계정 키 JSON (문자열)
  NOTION_TOKEN                  Notion 내부 통합 토큰
  GCHAT_WEBHOOK_URL              구글챗 수신 웹훅 URL

추가로 필요한 사전 준비 (v1 README의 서비스 계정 준비에 이어서):
  - 소재 원본 파일이 있는 구글 드라이브 폴더도 같은 서비스 계정에 "뷰어"로 공유
    (신규 소재 성과 탭의 소재명과 파일명이 정확히 일치하는 그 폴더)
  - 서비스 계정에 Google Drive API도 사용 설정 (Sheets API와 별개로 켜야 함)

사용법:
  python daily_report.py --date 2026-09-01
"""

import argparse
import datetime
import json
import os
import re
import sys
import time

import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build

SHEET_ID = "1iCeRn5-bPPFR47RydluurDcCBrKlYlMxwW02q5Bb9rA"
DAILY_TAB = "일별 Summary - Total"
CREATIVE_TAB = "신규 소재 성과"

# '신규 소재 성과' 탭 상단의 매체 필터 셀. 이 탭은 필터링된 뷰 하나만 제공해서
# 매체별로 보려면 이 셀 값을 실제로 바꿔써야 한다 — 데이터 검증 규칙에 있는
# 유효값만 순회하고, 끝나면 반드시 원래 값으로 되돌린다(공유 시트라 다른
# 사람 화면에도 순간적으로 영향을 줄 수 있음).
CREATIVE_MEDIA_FILTER_CELL = f"'{CREATIVE_TAB}'!Q5"
CREATIVE_MEDIA_OPTIONS = ["META", "Tiktok", "Moloco"]

NOTION_API = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"
DAILY_DB_ID = "3e6fbc1a-b2c3-8102-8e86-c4fae95fe55b"  # 일간 리포트 데이터베이스 ID (새 페이지 위치)

WEEKDAY_KR = ["월", "화", "수", "목", "금", "토", "일"]


# ── 인증 ──────────────────────────────────────────────────────────────
def google_creds():
    raw = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    info = json.loads(raw)
    return service_account.Credentials.from_service_account_info(
        info,
        scopes=[
            # 매체별 소재 효율(5번 섹션)을 보려면 필터 셀을 써야 해서 readonly가 아닌 전체 권한이 필요하다.
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.readonly",
        ],
    )


def sheets_client(creds):
    return build("sheets", "v4", credentials=creds)


def drive_client(creds):
    return build("drive", "v3", credentials=creds)


# ── 공통 유틸 ─────────────────────────────────────────────────────────
def parse_number(cell):
    if cell is None:
        return None
    s = str(cell).strip().replace(",", "")
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
    return f"{d.month}/{d.day} ({WEEKDAY_KR[d.weekday()]})"


def find_header_row(values, must_contain):
    for i, row in enumerate(values):
        if row and any(must_contain in (c or "") for c in row):
            return i
    raise RuntimeError(f"'{must_contain}' 포함 헤더 행을 못 찾았습니다.")


# ── 1. 일별 Summary - Total (오늘/전일/전주 동요일) ─────────────────────
COLUMN_MAP = {
    "집행 금액": "집행 금액(원)",
    "예산 소진율": "예산 소진율(%)",
    "앱설치 (Total)": "앱설치",
    "회원가입 (Total)": "회원가입",
    "설치당 단가 (Total)": "CPI(원)",
    "회원가입당 단가 (Total)": "CPA(원)",
    "가입 전환율 (Total)": "가입전환율(%)",
}


def fetch_daily_rows(service, target_date):
    resp = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=SHEET_ID, range=f"'{DAILY_TAB}'!A1:BN2000")
        .execute()
    )
    values = resp.get("values", [])

    # A열은 빈 스페이서 열이라 '일' 은 row[0]이 아니라 다른 열에 있다.
    header_idx, label_col = None, None
    for i, row in enumerate(values):
        if any("예산" in c for c in row):
            for j, cell in enumerate(row):
                if cell.strip() == "일":
                    header_idx, label_col = i, j
                    break
        if header_idx is not None:
            break
    if header_idx is None:
        raise RuntimeError("일별 Summary 헤더 행을 못 찾았습니다.")
    header = values[header_idx]

    # 'Total' 요약 섹션 뒤로 매체별 섹션이 같은 헤더 문구로 반복되므로,
    # label_col 다음의 첫 빈 칸까지만 잘라서 dict(zip(...))에서 중복 헤더가
    # 뒤 섹션 값으로 덮어써지는 걸 막는다.
    section_end = len(header)
    for j in range(label_col + 1, len(header)):
        if header[j].strip() == "":
            section_end = j
            break
    section_header = [h.strip() for h in header[label_col:section_end]]

    wanted = {
        "today": row_date_label(target_date),
        "prev_day": row_date_label(target_date - datetime.timedelta(days=1)),
        "prev_week_same_weekday": row_date_label(target_date - datetime.timedelta(days=7)),
    }
    found = {}
    for row in values[header_idx + 1 :]:
        if len(row) <= label_col:
            continue
        label = row[label_col].strip()
        for key, date_label in wanted.items():
            if label == date_label and key not in found:
                padded = row + [""] * (len(header) - len(row))
                found[key] = dict(zip(section_header, padded[label_col:section_end]))
    missing = [wanted[k] for k in wanted if k not in found]
    if missing:
        raise RuntimeError(f"일별 Summary에서 다음 날짜 행을 못 찾았습니다: {missing}")
    return found


# ── 2. MediaMix_YY.MM. (매체별 예산 배분) ───────────────────────────────
def fetch_mediamix(service, target_date):
    yy = target_date.strftime("%y")
    mm = target_date.strftime("%m")
    tab = f"MediaMix_{yy}.{mm}."
    try:
        resp = (
            service.spreadsheets()
            .values()
            # 헤더가 '예상 CPA' 열까지 이어져서 N열보다 오른쪽에 있다 — AB열까지 넉넉히 잡는다.
            .get(spreadsheetId=SHEET_ID, range=f"'{tab}'!A1:AB60")
            .execute()
        )
    except Exception as exc:  # noqa: BLE001
        print(f"[경고] {tab} 탭을 못 읽었습니다 ({exc}) — 매체별 효율 섹션은 비워둡니다.")
        return []
    values = resp.get("values", [])
    try:
        # '매체'로 찾으면 그 위의 섹션 제목 '매체 개요'에 먼저 걸린다.
        # '상품'은 실제 헤더 행에만 있는 고유한 키워드라 안전하다.
        header_idx = find_header_row(values, "상품")
    except RuntimeError:
        return []
    header = [h.strip() for h in values[header_idx]]

    def col(name):
        for i, h in enumerate(header):
            if name in h:
                return i
        return None

    def col2(*must_contain):
        """헤더 셀 안의 줄바꿈을 무시하고, 여러 키워드가 모두 들어있는 첫 열을 찾는다.
        예: col2("CPI", "TTL")은 '예상 CPI (paid)'는 건너뛰고 '예상 CPI (TTL)'을 찾는다."""
        for i, h in enumerate(header):
            norm = h.replace("\n", " ")
            if all(m in norm for m in must_contain):
                return i
        return None

    idx_media = col("매체")
    idx_product = col("상품")
    idx_budget = col("광고비")  # '예산'은 '예산 비중'(%) 열에도 걸려서 실제 금액 열인 '광고비'로 지정
    idx_share = col("비중")
    idx_cpi = col2("CPI", "TTL")
    idx_cpa = col2("CPA", "TTL")
    idx_install = col2("Install", "TTL")
    idx_signup = col2("회원가입", "TTL")

    rows = []
    last_media = ""
    for row in values[header_idx + 1 :]:
        if idx_product is None or len(row) <= idx_product:
            continue
        get = lambda i: row[i].strip() if i is not None and i < len(row) else ""
        product = get(idx_product)
        if not product:
            # 소계/총계 행(상품 칸이 비어있음)은 건너뛴다.
            continue
        # 매체 셀은 같은 매체의 상품이 여러 줄일 때 병합돼 있어, 빈 칸이면 위 행의 매체를 이어받는다.
        media = get(idx_media) or last_media
        last_media = media
        rows.append(
            {
                "매체": media,
                "상품": product,
                "예산": parse_number(get(idx_budget)),
                "비중": parse_number(get(idx_share)),
                "install": parse_number(get(idx_install)),
                "signup": parse_number(get(idx_signup)),
            }
        )
    return rows


def aggregate_mediamix(rows):
    """상품별(AOS/iOS 등)로 나뉜 행을 매체 하나당 한 행으로 합친다.
    CPI/CPA는 상품별 값을 평균 내는 대신, 예산/Install/가입을 합산한 뒤
    다시 나눠서 블렌디드 값을 계산한다."""
    order = []
    by_media = {}
    for r in rows:
        media = r["매체"]
        if media not in by_media:
            by_media[media] = {"products": [], "예산": 0, "install": 0, "signup": 0}
            order.append(media)
        agg = by_media[media]
        agg["products"].append(r["상품"])
        agg["예산"] += r["예산"] or 0
        agg["install"] += r["install"] or 0
        agg["signup"] += r["signup"] or 0

    total_budget = sum(agg["예산"] for agg in by_media.values()) or None

    def label(media, products):
        if len(products) <= 1:
            return media
        suffixes = [m.group(1) for p in products if (m := re.search(r"\(([^)]+)\)", p))]
        return f"{media} ({'/'.join(suffixes)})" if suffixes else media

    result = []
    for media in order:
        agg = by_media[media]
        cpi = round(agg["예산"] / agg["install"]) if agg["install"] else None
        cpa = round(agg["예산"] / agg["signup"]) if agg["signup"] else None
        share = (agg["예산"] / total_budget) if total_budget else None
        result.append(
            {
                "매체": label(media, agg["products"]),
                "예산": agg["예산"],
                "비중": share,
                "CPI": cpi,
                "CPA": cpa,
            }
        )
    return result


# ── 3. 신규 소재 성과 (당일 운영 중인 소재) ──────────────────────────────
def fetch_active_creatives(service, target_date):
    resp = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=SHEET_ID, range=f"'{CREATIVE_TAB}'!A1:T3000")
        .execute()
    )
    values = resp.get("values", [])
    header_idx = find_header_row(values, "소재 운영 시작")
    header = [h.strip() for h in values[header_idx]]

    def col(name):
        for i, h in enumerate(header):
            if name == h or name in h:
                return i
        return None

    idx = {
        "start": col("소재 운영 시작"),
        "end": col("D+14"),
        "name": col("소재명"),
        "usp": col("USP"),
        "influencer": col("인플루언서"),
        "spend": col("집행 금액"),
        "install": col("설치"),
        "cpi": col("CPI") or col("설치당"),
        "signup": col("회원가입") or col("가입"),
        "cpa": col("CPA") or col("회원가입당"),
        "cvr": col("가입 전환율") or col("가입전환율"),
    }

    def parse_date(s):
        try:
            return datetime.date.fromisoformat(s.strip())
        except Exception:  # noqa: BLE001
            return None

    active = []
    for row in values[header_idx + 1 :]:
        if not row:
            continue
        get = lambda key: row[idx[key]] if idx.get(key) is not None and idx[key] < len(row) else ""
        start = parse_date(get("start"))
        end = parse_date(get("end"))
        if not start or not end:
            continue
        if not (start <= target_date <= end):
            continue
        name = get("name").strip()
        if not name:
            continue
        active.append(
            {
                "name": name,
                "usp": get("usp").strip(),
                "influencer": get("influencer").strip(),
                "period": f"{start.month}/{start.day}~{end.month}/{end.day}",
                "cpa": parse_number(get("cpa")),
                "cvr": parse_number(get("cvr")),
                "install": parse_number(get("install")),
                "signup": parse_number(get("signup")),
                "spend": parse_number(get("spend")),
            }
        )
    return active


def rank_creatives(active):
    with_signup = [c for c in active if (c["signup"] or 0) > 0]
    zero_signup = [c for c in active if not c["signup"]]
    high = sorted(with_signup, key=lambda c: c["cpa"] or float("inf"))[:3]
    low = sorted(zero_signup, key=lambda c: -(c["spend"] or 0))[:3]
    return high, low


# ── 3-1. 매체별 소재 효율 (필터 셀을 순회하며 읽기) ──────────────────────
def read_creative_media_filter(service):
    resp = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=CREATIVE_MEDIA_FILTER_CELL
    ).execute()
    values = resp.get("values", [])
    return values[0][0] if values and values[0] else "ALL"


def write_creative_media_filter(service, media):
    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=CREATIVE_MEDIA_FILTER_CELL,
        valueInputOption="RAW",
        body={"values": [[media]]},
    ).execute()
    time.sleep(1.5)  # 필터에 딸린 수식이 재계산될 시간을 준다.


def fetch_creatives_by_media(service, target_date):
    """CREATIVE_MEDIA_OPTIONS를 순회하며 필터 셀을 바꿔쓰고 그때마다 활성 소재를
    다시 읽는다. 공유 시트의 필터를 실제로 바꾸는 것이므로, 중간에 실패해도
    finally에서 원래 값으로 반드시 복원한다."""
    original = read_creative_media_filter(service)
    result = {}
    try:
        for media in CREATIVE_MEDIA_OPTIONS:
            write_creative_media_filter(service, media)
            result[media] = fetch_active_creatives(service, target_date)
    finally:
        write_creative_media_filter(service, original)
    return result


def summarize_media_creatives(creatives):
    if not creatives:
        return "해당 날짜에 활성 소재 없음"
    with_signup = [c for c in creatives if (c["signup"] or 0) > 0]
    parts = [f"소재 {len(creatives)}개"]
    if with_signup:
        total_spend = sum(c["spend"] or 0 for c in with_signup)
        total_signup = sum(c["signup"] or 0 for c in with_signup)
        if total_signup:
            parts.append(f"평균 CPA {round(total_spend / total_signup):,}원")
        top = min(with_signup, key=lambda c: c["cpa"] or float("inf"))
        parts.append(f"최고효율 {top['name']}(CPA {top['cpa']:,}원)")
    else:
        parts.append("가입 전환 소재 없음")
    return " · ".join(parts)


# ── 4. 구글 드라이브에서 소재 원본 링크 찾기 ─────────────────────────────
def drive_link(drive, creative_name):
    escaped = creative_name.replace("'", "\\'")
    q = f"name contains '{escaped}' and trashed = false"
    resp = drive.files().list(q=q, fields="files(id, webViewLink)", pageSize=1).execute()
    files = resp.get("files", [])
    if not files:
        return None
    return files[0].get("webViewLink")


# ── Notion 블록 빌더 ─────────────────────────────────────────────────
def h2(text):
    return {"object": "block", "type": "heading_2", "heading_2": {"rich_text": [{"type": "text", "text": {"content": text}}]}}


def para(text):
    return {"object": "block", "type": "paragraph", "paragraph": {"rich_text": [{"type": "text", "text": {"content": text}}]}}


def bullet(text):
    return {
        "object": "block",
        "type": "bulleted_list_item",
        "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": text}}]},
    }


def numbered(text):
    return {
        "object": "block",
        "type": "numbered_list_item",
        "numbered_list_item": {"rich_text": [{"type": "text", "text": {"content": text}}]},
    }


def table(headers, rows, link_col=None):
    def cell(text, url=None):
        t = {"type": "text", "text": {"content": str(text)}}
        if url:
            t["text"]["link"] = {"url": url}
        return [t]

    table_rows = [
        {"object": "block", "type": "table_row", "table_row": {"cells": [cell(h) for h in headers]}}
    ]
    for r in rows:
        cells = []
        for i, val in enumerate(r):
            url = val[1] if link_col == i and isinstance(val, tuple) else None
            text = val[0] if isinstance(val, tuple) else val
            cells.append(cell(text, url))
        table_rows.append({"object": "block", "type": "table_row", "table_row": {"cells": cells}})
    return {
        "object": "block",
        "type": "table",
        "table": {
            "table_width": len(headers),
            "has_column_header": True,
            "has_row_header": False,
            "children": table_rows,
        },
    }


def build_children(target_date, daily, mediamix, high, low, media_creatives):
    def g(row, col):
        return parse_number(row.get(col))

    children = [h2("1. 전체 예산 / 집행 금액 / 소진율")]
    children.append(
        para(f"{target_date.isoformat()} 집행 {g(daily['today'], '집행 금액') or '-'}원 / "
             f"예산 소진율 {g(daily['today'], '예산 소진율') or '-'}")
    )

    children.append(h2("2. 전체 효율"))
    install_t = g(daily["today"], "앱설치 (Total)")
    signup_t = g(daily["today"], "회원가입 (Total)")
    install_p = g(daily["prev_day"], "앱설치 (Total)")
    signup_p = g(daily["prev_day"], "회원가입 (Total)")
    install_w = g(daily["prev_week_same_weekday"], "앱설치 (Total)")
    signup_w = g(daily["prev_week_same_weekday"], "회원가입 (Total)")
    pct = lambda new, old: f"{round((new - old) / old * 100, 1)}%" if new is not None and old else "-"
    children.append(
        table(
            ["지표", target_date.isoformat(), "전일 대비", "전주 동요일 대비"],
            [
                ["앱설치", install_t, pct(install_t, install_p), pct(install_t, install_w)],
                ["회원가입", signup_t, pct(signup_t, signup_p), pct(signup_t, signup_w)],
                ["CPI", g(daily["today"], "설치당 단가 (Total)"), "-", "-"],
                ["CPA", g(daily["today"], "회원가입당 단가 (Total)"), "-", "-"],
                ["가입전환율", g(daily["today"], "가입 전환율 (Total)"), "-", "-"],
            ],
        )
    )

    children.append(h2("3. 매체별 효율"))
    if mediamix:
        children.append(para("9월 매체별 예산 배분(계획). 매체별 실적은 원본 시트에 없어 계획 대비 비교는 별도 연동이 필요합니다."))
        children.append(
            table(
                ["매체", "예산", "비중", "예상 CPI", "예상 CPA"],
                [
                    [
                        r["매체"],
                        f"{r['예산']:,}" if r["예산"] is not None else "-",
                        f"{r['비중'] * 100:.1f}%" if r["비중"] is not None else "-",
                        f"{r['CPI']:,}" if r["CPI"] is not None else "-",
                        f"{r['CPA']:,}" if r["CPA"] is not None else "-",
                    ]
                    for r in mediamix
                ],
            )
        )
    else:
        children.append(para("이번 달 MediaMix 탭을 찾지 못했습니다 — 탭 이름 규칙이 바뀌었을 수 있습니다."))

    children.append(h2("4. 전체 소재 효율"))
    children.append(para(f"{target_date.isoformat()} 기준 운영 중인 소재 (D+14 집계 기준 수치)."))
    if high:
        children.append(
            table(
                ["소재", "운영기간", "USP", "CPA", "가입전환율"],
                [
                    [(c["name"], c.get("link")) if c.get("link") else c["name"], c["period"], c["usp"], c["cpa"], c["cvr"]]
                    for c in high
                ],
                link_col=0,
            )
        )
    if low:
        children.append(para("저효율 (중단/교체 후보)"))
        children.append(
            table(
                ["소재", "운영기간", "집행 금액", "설치/가입"],
                [
                    [(c["name"], c.get("link")) if c.get("link") else c["name"], c["period"], c["spend"], f"{c['install'] or 0}/{c['signup'] or 0}"]
                    for c in low
                ],
                link_col=0,
            )
        )
    if not high and not low:
        children.append(para("해당 날짜에 활성 소재 데이터를 찾지 못했습니다."))

    children.append(h2("5. 매체별 소재 효율"))
    if media_creatives:
        children.append(
            para(
                "'신규 소재 성과' 탭의 매체 필터를 META/Tiktok/Moloco로 순회하며 읽었습니다 "
                "(Google UAC·ASA·네이버BSA는 이 탭에 소재 단위 추적이 없어 필터 옵션에도 없습니다)."
            )
        )
        children.append(
            table(
                ["매체", "소재별 효율"],
                [[media, summarize_media_creatives(cs)] for media, cs in media_creatives.items()],
            )
        )
    else:
        children.append(para("매체별 소재 효율을 읽지 못했습니다 — 필터 셀 접근에 실패했을 수 있습니다."))

    children.append(h2("6. 개선점"))
    improvements = []
    if low:
        names = ", ".join(c["name"] for c in low)
        improvements.append(f"저효율 소재 정리 검토 대상: {names}")
    if high:
        top = high[0]
        improvements.append(f"고효율 패턴 확장 후보: {top['name']} (CPA {top['cpa']}원, USP: {top['usp']})")
    improvements.append("매체별 실적 데이터가 없어 채널 간 예산 재배분 판단이 어렵습니다 — 매체 API 연동을 우선 검토하세요.")
    for text in improvements:
        children.append(numbered(text))

    return children


def build_properties(target_date, daily):
    def g(row, col):
        return parse_number(row.get(col))

    install_t = g(daily["today"], "앱설치 (Total)")
    signup_t = g(daily["today"], "회원가입 (Total)")
    install_p = g(daily["prev_day"], "앱설치 (Total)")
    signup_p = g(daily["prev_day"], "회원가입 (Total)")
    install_w = g(daily["prev_week_same_weekday"], "앱설치 (Total)")
    signup_w = g(daily["prev_week_same_weekday"], "회원가입 (Total)")
    pct = lambda new, old: round((new - old) / old, 4) if new is not None and old else None

    return {
        "날짜": {"title": [{"text": {"content": target_date.isoformat()}}]},
        "요일": {"select": {"name": WEEKDAY_KR[target_date.weekday()]}},
        "집행 금액(원)": {"number": g(daily["today"], "집행 금액")},
        "예산 소진율(%)": {"number": g(daily["today"], "예산 소진율")},
        "앱설치": {"number": install_t},
        "회원가입": {"number": signup_t},
        "CPI(원)": {"number": g(daily["today"], "설치당 단가 (Total)")},
        "CPA(원)": {"number": g(daily["today"], "회원가입당 단가 (Total)")},
        "가입전환율(%)": {"number": g(daily["today"], "가입 전환율 (Total)")},
        "전일 대비 설치 증감(%)": {"number": pct(install_t, install_p)},
        "전일 대비 가입 증감(%)": {"number": pct(signup_t, signup_p)},
        "전주 동요일 대비 설치 증감(%)": {"number": pct(install_t, install_w)},
        "전주 동요일 대비 가입 증감(%)": {"number": pct(signup_t, signup_w)},
    }


def create_notion_page(properties, children):
    headers = {
        "Authorization": f"Bearer {os.environ['NOTION_TOKEN']}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }
    payload = {"parent": {"database_id": DAILY_DB_ID}, "properties": properties, "children": children}
    resp = requests.post(f"{NOTION_API}/pages", headers=headers, json=payload, timeout=30)
    if not resp.ok:
        print(resp.text, file=sys.stderr)
    resp.raise_for_status()
    return resp.json()


def notify_gchat(page_url, target_date):
    webhook = os.environ.get("GCHAT_WEBHOOK_URL")
    if not webhook:
        print("GCHAT_WEBHOOK_URL 미설정 — 알림 생략")
        return
    text = f"📊 {target_date.isoformat()} 파스타 일간 리포트가 생성됐습니다.\n{page_url}"
    resp = requests.post(webhook, json={"text": text}, timeout=15)
    resp.raise_for_status()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date")
    args = parser.parse_args()
    target_date = datetime.date.fromisoformat(args.date) if args.date else datetime.date.today()

    creds = google_creds()
    sheets = sheets_client(creds)
    drive = drive_client(creds)

    daily = fetch_daily_rows(sheets, target_date)
    mediamix = aggregate_mediamix(fetch_mediamix(sheets, target_date))
    active = fetch_active_creatives(sheets, target_date)
    drive_ok = True
    for c in active:
        if not drive_ok:
            c["link"] = None
            continue
        try:
            c["link"] = drive_link(drive, c["name"])
        except Exception as exc:  # noqa: BLE001
            print(f"[경고] 드라이브 조회 실패 ({exc}) — 소재 링크는 비워둡니다.")
            drive_ok = False
            c["link"] = None
    high, low = rank_creatives(active)

    try:
        media_creatives = fetch_creatives_by_media(sheets, target_date)
    except Exception as exc:  # noqa: BLE001
        print(f"[경고] 매체별 소재 필터 순회 실패 ({exc}) — 5번 섹션은 비워둡니다.")
        media_creatives = {}

    properties = build_properties(target_date, daily)
    children = build_children(target_date, daily, mediamix, high, low, media_creatives)

    page = create_notion_page(properties, children)
    page_url = page.get("url", "")
    print(f"노션 페이지 생성 완료: {page_url}")
    notify_gchat(page_url, target_date)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"실패: {exc}", file=sys.stderr)
        sys.exit(1)

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
  python daily_report.py --date 2026-09-01   # 특정 날짜 하나만 (백필/테스트용)
  python daily_report.py                     # 자동 모드: 지금(KST) 기준으로 대행사
                                              # 데이터가 이미 올라왔지만 아직 노션에
                                              # 발행 안 된 날짜를 전부 찾아 발행한다.

자동 모드가 쓰는 업데이트 규칙 (README 참고):
  - 영업일(평일이면서 대한민국 공휴일이 아닌 날) 데이터: 그 다음 영업일 오후 1시(KST)
  - 주말·공휴일 데이터: 연휴가 끝난 다음 영업일 오후 1시(KST)에 한꺼번에 올라온다
    (예: 금/토/일 → 월요일, 추석 연휴 3일 → 연휴 다음 영업일)
"""

import argparse
import datetime
import json
import os
import re
import sys
import time

import holidays
import requests
from google.oauth2 import service_account

import charts
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

KST = datetime.timezone(datetime.timedelta(hours=9))
AUTO_LOOKBACK_DAYS = 14  # 자동 모드에서 "발행됐어야 하는데 빠졌는지" 되돌아볼 기간

NOTION_API = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"
DAILY_DB_ID = "3e6fbc1a-b2c3-8102-8e86-c4fae95fe55b"  # 일간 리포트 데이터베이스 ID (새 페이지 위치)

META_API = "https://graph.facebook.com/v21.0"
META_AD_ACCOUNT_ID = "act_343984491884470"  # 파스타 광고 계정
AGE_BUCKET_ORDER = ["13-17", "18-24", "25-34", "35-44", "45-54", "55-64", "65+"]

AIRBRIDGE_API = "https://api.airbridge.io"
AIRBRIDGE_APP_NAME = "pasta"
# Airbridge의 channel 값 -> 리포트에서 쓰는 매체 표기. 시트(MediaMix)에는 매체별
# 실적이 없어서 UAC/네이버BSA는 지금까지 "소재 단위 추적 없음"으로만 표시했는데,
# Airbridge(MMP)는 모든 매체의 실제 설치·가입을 채널 단위로 집계해주기 때문에
# 이걸로 그 공백을 메운다.
AIRBRIDGE_CHANNEL_LABEL = {
    "google.adwords": "UAC",
    "facebook.business": "META",
    "moloco": "Moloco",
    "apple.searchads": "Apple Search Ads",
    "bsa": "네이버BSA",
}

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


# ── 자동 발행: 이 날짜의 데이터가 지금(KST) 이미 올라와 있는지 판단 ──────────
# 대한민국 공휴일(설날/추석 같은 음력 연휴 포함)도 대행사가 쉬는 날이라, 주말과
# 똑같이 "영업일이 아닌 날"로 취급한다. holidays 라이브러리가 연도별로 계산해준다.
KR_HOLIDAYS = holidays.KR(years=range(2020, 2036))


def is_business_day(d: datetime.date) -> bool:
    return d.weekday() < 5 and d not in KR_HOLIDAYS


def next_business_day(d: datetime.date) -> datetime.date:
    d = d + datetime.timedelta(days=1)
    while not is_business_day(d):
        d += datetime.timedelta(days=1)
    return d


def data_ready_at(d: datetime.date) -> datetime.datetime:
    """대행사가 날짜 d의 데이터를 올리는 시점(KST)을 돌려준다.
    영업일(평일이면서 공휴일이 아닌 날) 데이터는 그 다음 영업일 오후 1시에 올라온다.
    주말·공휴일이 여러 날 이어져도(연휴), 그 사이에 낀 날짜들은 모두 연휴가 끝난
    다음 영업일에 한꺼번에 올라온다 — 금/토/일이 월요일에 함께 올라오는 것과 같은
    원리를 공휴일에도 그대로 적용한 것."""
    ready_date = next_business_day(d)
    return datetime.datetime.combine(ready_date, datetime.time(13, 0), tzinfo=KST)


def is_data_ready(d: datetime.date, as_of: datetime.datetime) -> bool:
    return as_of >= data_ready_at(d)


def pending_dates(as_of: datetime.datetime, published: set, lookback_days: int = AUTO_LOOKBACK_DAYS):
    """as_of(KST) 시점에 데이터가 이미 준비됐지만 아직 published에 없는 날짜들을,
    오래된 날짜부터 순서대로 돌려준다."""
    today = as_of.date()
    result = []
    for i in range(lookback_days, 0, -1):
        d = today - datetime.timedelta(days=i)
        if d in published:
            continue
        if is_data_ready(d, as_of):
            result.append(d)
    return result


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


def fetch_daily_tab_values(service):
    resp = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=SHEET_ID, range=f"'{DAILY_TAB}'!A1:BN2000")
        .execute()
    )
    return resp.get("values", [])


def locate_daily_header(values):
    """'일별 Summary - Total' 탭의 헤더 행 인덱스, '일' 레이블 열, 그 구간의
    헤더 목록을 찾는다. A열은 빈 스페이서 열이라 '일' 은 row[0]이 아니라 다른
    열에 있고, 'Total' 요약 섹션 뒤로 매체별 섹션이 같은 헤더 문구로
    반복되므로 label_col 다음의 첫 빈 칸까지만 잘라서 중복 헤더를 피한다."""
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

    section_end = len(header)
    for j in range(label_col + 1, len(header)):
        if header[j].strip() == "":
            section_end = j
            break
    section_header = [h.strip() for h in header[label_col:section_end]]
    return header_idx, label_col, section_end, section_header


def fetch_daily_rows(service, target_date):
    values = fetch_daily_tab_values(service)
    header_idx, label_col, section_end, section_header = locate_daily_header(values)
    header = values[header_idx]

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


def fetch_mtd_summary(service, target_date):
    """이번 달 1일부터 target_date까지의 일별 행을 모두 모아 누적 지표를
    직접 계산한다. 시트에 있는 월별 합계 행은 시트 자체의 '오늘' 셀을 기준으로
    계산돼 있어서, 과거 날짜로 백필 테스트하면 그 날짜가 아니라 실제 오늘까지
    누적된 값이 나올 수 있다 — 그래서 매번 일별 행을 직접 합산한다."""
    values = fetch_daily_tab_values(service)
    header_idx, label_col, section_end, section_header = locate_daily_header(values)
    header = values[header_idx]
    month_start = target_date.replace(day=1)

    def g(row, col):
        return parse_number(row.get(col))

    total_spend = total_install = total_signup = 0
    days_counted = 0
    for row in values[header_idx + 1 :]:
        if len(row) <= label_col:
            continue
        m = re.match(r"^(\d{1,2})/(\d{1,2}) \(", row[label_col].strip())
        if not m:
            continue
        month, day = int(m.group(1)), int(m.group(2))
        try:
            d = datetime.date(target_date.year, month, day)
        except ValueError:
            continue
        if not (month_start <= d <= target_date):
            continue
        padded = row + [""] * (len(header) - len(row))
        r = dict(zip(section_header, padded[label_col:section_end]))
        total_spend += g(r, "집행 금액") or 0
        total_install += g(r, "앱설치 (Total)") or 0
        total_signup += g(r, "회원가입 (Total)") or 0
        days_counted += 1

    return {
        "days": days_counted,
        "spend": total_spend,
        "install": total_install,
        "signup": total_signup,
        "cpi": round(total_spend / total_install) if total_install else None,
        "cpa": round(total_spend / total_signup) if total_signup else None,
        "cvr": round(total_signup / total_install, 4) if total_install else None,
    }


def fetch_mtd_series(service, target_date):
    """fetch_mtd_summary와 같은 파싱을 쓰지만, 합계 대신 날짜별 값을 그대로
    리스트로 돌려준다 — 그래프 그릴 때 쓴다."""
    values = fetch_daily_tab_values(service)
    header_idx, label_col, section_end, section_header = locate_daily_header(values)
    header = values[header_idx]
    month_start = target_date.replace(day=1)

    def g(row, col):
        return parse_number(row.get(col))

    series = []
    for row in values[header_idx + 1 :]:
        if len(row) <= label_col:
            continue
        m = re.match(r"^(\d{1,2})/(\d{1,2}) \(", row[label_col].strip())
        if not m:
            continue
        month, day = int(m.group(1)), int(m.group(2))
        try:
            d = datetime.date(target_date.year, month, day)
        except ValueError:
            continue
        if not (month_start <= d <= target_date):
            continue
        padded = row + [""] * (len(header) - len(row))
        r = dict(zip(section_header, padded[label_col:section_end]))
        series.append(
            {
                "date": d,
                "install": g(r, "앱설치 (Total)"),
                "signup": g(r, "회원가입 (Total)"),
                "spend": g(r, "집행 금액"),
            }
        )
    series.sort(key=lambda x: x["date"])
    return series


GENDER_LABEL_KR = {"female": "여성", "male": "남성", "unknown": "성별 미확인"}


def fetch_meta_signup_age_gender(target_date):
    """META 광고 계정에서 이번 달(1일~target_date) 회원가입(앱 SDK
    complete_registration 이벤트) 건수를 연령대 x 성별로 집계한다. 다른
    매체는 연령 데이터를 아예 안 줘서, 이건 META 전용 지표로만 취급해야 한다."""
    token = os.environ["META_ACCESS_TOKEN"]
    month_start = target_date.replace(day=1)
    params = {
        "access_token": token,
        "level": "account",
        "fields": "actions",
        "breakdowns": "age,gender",
        "action_breakdowns": "action_type",
        "time_range": json.dumps({"since": month_start.isoformat(), "until": target_date.isoformat()}),
        "limit": 100,
    }
    resp = requests.get(f"{META_API}/{META_AD_ACCOUNT_ID}/insights", params=params, timeout=30)
    resp.raise_for_status()
    rows = resp.json().get("data", [])

    by_age_gender = {age: {"female": 0, "male": 0, "unknown": 0} for age in AGE_BUCKET_ORDER}
    for row in rows:
        age = row.get("age")
        gender = row.get("gender")
        if age not in by_age_gender or gender not in ("female", "male", "unknown"):
            continue
        for action in row.get("actions", []):
            if action["action_type"] == "app_custom_event.fb_mobile_complete_registration":
                by_age_gender[age][gender] += int(float(action["value"]))
                break
    return by_age_gender


def fetch_airbridge_channel_actuals(date_from, date_to, timeout_sec=20):
    """Airbridge(MMP)에서 date_from~date_to 기간의 채널별 실제 설치·가입 수를
    가져온다. 시트의 MediaMix 탭은 예산 '계획'만 있고 실적이 없어서,
    UAC·네이버BSA처럼 소재 단위 추적이 없는 매체도 이걸로 실제 성과를
    보여줄 수 있다. 쿼리는 비동기라 taskId를 받고 SUCCESS 될 때까지 잠깐
    polling한다."""
    token = os.environ["AIRBRIDGE_API_TOKEN"]
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "from": date_from.isoformat(),
        "to": date_to.isoformat(),
        "groupBys": ["channel"],
        "metrics": ["app_installs", "app_sign_up"],
    }
    resp = requests.post(
        f"{AIRBRIDGE_API}/reports/api/v7/apps/{AIRBRIDGE_APP_NAME}/actuals/query",
        headers=headers, json=body, timeout=30,
    )
    resp.raise_for_status()
    task_id = resp.json()["task"]["taskId"]

    deadline = time.time() + timeout_sec
    result_url = f"{AIRBRIDGE_API}/reports/api/v7/apps/{AIRBRIDGE_APP_NAME}/actuals/query/{task_id}"
    while True:
        resp = requests.get(result_url, headers=headers, params={"skip": 0, "size": 100}, timeout=30)
        resp.raise_for_status()
        payload = resp.json()
        status = payload["task"]["status"]
        if status == "SUCCESS":
            break
        if status in ("FAILURE", "CANCELED"):
            raise RuntimeError(f"Airbridge 쿼리 실패 (status={status})")
        if time.time() > deadline:
            raise TimeoutError("Airbridge 쿼리가 시간 내에 끝나지 않았습니다")
        time.sleep(1.5)

    by_channel = {}
    for row in payload["actuals"]["data"]["rows"]:
        channel = row["groupBys"][0]
        label = AIRBRIDGE_CHANNEL_LABEL.get(channel)
        if not label:
            continue
        install = row["values"].get("app_installs", {}).get("value")
        signup = row["values"].get("app_sign_up", {}).get("value")
        by_channel[label] = {
            "install": int(install) if install is not None else None,
            "signup": int(signup) if signup is not None else None,
        }
    return by_channel


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
def drive_file_info(drive, creative_name):
    """소재 원본 파일의 열람 링크와 썸네일 URL을 한 번의 조회로 가져온다.
    thumbnailLink는 인증 없이도 바로 이미지가 뜨는 URL이라(직접 테스트로 확인),
    노션에 외부 이미지로 그대로 박아넣을 수 있다."""
    escaped = creative_name.replace("'", "\\'")
    q = f"name contains '{escaped}' and trashed = false"
    resp = drive.files().list(
        q=q,
        fields="files(id, webViewLink, thumbnailLink)",
        pageSize=1,
        # 소재 원본이 '내 드라이브'가 아니라 공유 드라이브(Shared Drive)에 있어서,
        # 이 옵션이 없으면 서비스 계정에 아무 파일도 안 보인다.
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora="allDrives",
    ).execute()
    files = resp.get("files", [])
    if not files:
        return {"link": None, "thumbnail": None}
    return {"link": files[0].get("webViewLink"), "thumbnail": files[0].get("thumbnailLink")}


def drive_link(drive, creative_name):
    return drive_file_info(drive, creative_name)["link"]


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


def thumb_block(name, url):
    return {
        "object": "block",
        "type": "image",
        "image": {
            "type": "external",
            "external": {"url": url},
            "caption": [{"type": "text", "text": {"content": name}}],
        },
    }


def callout(emoji, text):
    return {
        "object": "block",
        "type": "callout",
        "callout": {
            "rich_text": [{"type": "text", "text": {"content": text}}],
            "icon": {"type": "emoji", "emoji": emoji},
        },
    }


def date_with_weekday(d: datetime.date) -> str:
    return f"{d.isoformat()} ({WEEKDAY_KR[d.weekday()]})"


def fmt_num(value, unit=""):
    return "-" if value is None else f"{value:,}{unit}"


def fmt_ratio_pct(value):
    """0.641 같은 절대 비율을 '64.1%'로. 증감(+/-)이 아닌 값에 쓴다."""
    return "-" if value is None else f"{value * 100:.1f}%"


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


def find_media(mediamix, keyword):
    for r in mediamix:
        if keyword in r["매체"]:
            return r
    return None


def build_children(
    target_date, daily, mediamix, high, low, media_creatives, mtd,
    mtd_chart_id=None, age_chart_id=None, airbridge_actuals=None, media_chart_id=None,
):
    def g(row, col):
        return parse_number(row.get(col))

    date_label = date_with_weekday(target_date)
    children = [
        callout("🗓️", f"{date_label} 파스타 일간 리포트"),
        h2(f"📅 이번 달 누적 요약 ({target_date.year}-{target_date.month:02d}-01 ~ {target_date.isoformat()}, {mtd['days']}일)"),
    ]
    month_budget = sum(r["예산"] or 0 for r in mediamix) or None
    exhaustion = f"{mtd['spend'] / month_budget * 100:.1f}%" if month_budget else "-"
    children.append(
        para(
            f"💰 집행 {fmt_num(mtd['spend'], '원')}"
            + (f" / 이번 달 예산 {fmt_num(month_budget, '원')} 대비 {exhaustion} 소진" if month_budget else "")
        )
    )
    children.append(
        table(
            ["지표", "누적"],
            [
                ["앱설치", fmt_num(mtd["install"])],
                ["회원가입", fmt_num(mtd["signup"])],
                ["CPI", fmt_num(mtd["cpi"], "원")],
                ["CPA", fmt_num(mtd["cpa"], "원")],
                ["가입전환율", fmt_ratio_pct(mtd["cvr"])],
            ],
        )
    )
    if mtd_chart_id:
        children.append(image_upload_block(mtd_chart_id, "이번 달 일별 앱설치 · 회원가입 추이"))
    if age_chart_id:
        children.append(para("👥 META 광고 전환(회원가입) 기준 연령대·성별 분포 — 연령 데이터는 현재 META만 제공합니다."))
        children.append(image_upload_block(age_chart_id, "META 가입자 연령대·성별 분포 (이번 달 누적)"))
    if media_chart_id:
        children.append(para("📡 매체별 실제 설치 수 (Airbridge 기준, 이번 달 누적) — UAC·네이버BSA 포함 전체 매체."))
        children.append(image_upload_block(media_chart_id, "이번 달 매체별 실제 설치 수"))

    children.append(h2(f"1. 💰 전체 예산 / 집행 금액 / 소진율 — {date_label}"))
    children.append(
        para(
            f"집행 {fmt_num(g(daily['today'], '집행 금액'), '원')} / "
            f"예산 소진율 {fmt_ratio_pct(g(daily['today'], '예산 소진율'))}"
        )
    )

    children.append(h2(f"2. 📈 전체 효율 — {date_label}"))
    install_t = g(daily["today"], "앱설치 (Total)")
    signup_t = g(daily["today"], "회원가입 (Total)")
    install_p = g(daily["prev_day"], "앱설치 (Total)")
    signup_p = g(daily["prev_day"], "회원가입 (Total)")
    install_w = g(daily["prev_week_same_weekday"], "앱설치 (Total)")
    signup_w = g(daily["prev_week_same_weekday"], "회원가입 (Total)")
    pct = lambda new, old: f"{round((new - old) / old * 100, 1):+.1f}%" if new is not None and old else "-"
    children.append(
        table(
            ["지표", date_label, "전일 대비", "전주 동요일 대비"],
            [
                ["앱설치", fmt_num(install_t), pct(install_t, install_p), pct(install_t, install_w)],
                ["회원가입", fmt_num(signup_t), pct(signup_t, signup_p), pct(signup_t, signup_w)],
                ["CPI", fmt_num(g(daily["today"], "설치당 단가 (Total)"), "원"), "-", "-"],
                ["CPA", fmt_num(g(daily["today"], "회원가입당 단가 (Total)"), "원"), "-", "-"],
                ["가입전환율", fmt_ratio_pct(g(daily["today"], "가입 전환율 (Total)")), "-", "-"],
            ],
        )
    )

    children.append(h2("3. 📡 매체별 효율"))
    if mediamix:
        note = f"{target_date.month}월 매체별 예산 배분(계획) vs {date_label} 실제 성과(Airbridge 기준)."
        if not airbridge_actuals:
            note += " (Airbridge 연동 안 됨 — 실제 성과 칸은 비어 있습니다)"
        children.append(para(note))

        def match_actual(media_label):
            for key, v in (airbridge_actuals or {}).items():
                if key in media_label:
                    return v
            return None

        rows = []
        for r in mediamix:
            actual = match_actual(r["매체"])
            rows.append(
                [
                    r["매체"],
                    fmt_num(r["예산"], "원"),
                    fmt_ratio_pct(r["비중"]),
                    fmt_num(r["CPI"], "원"),
                    fmt_num(r["CPA"], "원"),
                    fmt_num(actual["install"]) if actual else "-",
                    fmt_num(actual["signup"]) if actual else "-",
                    fmt_ratio_pct(actual["signup"] / actual["install"]) if actual and actual["install"] else "-",
                ]
            )
        children.append(
            table(["매체", "예산(계획)", "비중", "예상 CPI", "예상 CPA", "실제 설치", "실제 가입", "실제 가입전환율"], rows)
        )
    else:
        children.append(para("이번 달 MediaMix 탭을 찾지 못했습니다 — 탭 이름 규칙이 바뀌었을 수 있습니다."))

    children.append(h2(f"4. 🎨 전체 소재 효율 — {date_label}"))
    children.append(para(f"{date_label} 기준 운영 중인 소재 (D+14 집계 기준 수치)."))
    if high:
        children.append(para("🟢 고효율"))
        children.append(
            table(
                ["소재", "운영기간", "USP", "CPA", "가입전환율"],
                [
                    [
                        (c["name"], c.get("link")) if c.get("link") else c["name"],
                        c["period"],
                        c["usp"],
                        fmt_num(c["cpa"], "원"),
                        fmt_ratio_pct(c["cvr"]),
                    ]
                    for c in high
                ],
                link_col=0,
            )
        )
        for c in high:
            if c.get("thumbnail"):
                children.append(thumb_block(c["name"], c["thumbnail"]))
    if low:
        children.append(para("🔴 저효율 (중단/교체 후보)"))
        children.append(
            table(
                ["소재", "운영기간", "집행 금액", "설치/가입"],
                [
                    [
                        (c["name"], c.get("link")) if c.get("link") else c["name"],
                        c["period"],
                        fmt_num(c["spend"], "원"),
                        f"{c['install'] or 0}/{c['signup'] or 0}",
                    ]
                    for c in low
                ],
                link_col=0,
            )
        )
        for c in low:
            if c.get("thumbnail"):
                children.append(thumb_block(c["name"], c["thumbnail"]))
    if not high and not low:
        children.append(para("해당 날짜에 활성 소재 데이터를 찾지 못했습니다."))

    children.append(h2("5. 📱 매체별 소재 효율"))
    rows = []
    if media_creatives:
        rows.extend([media, summarize_media_creatives(cs)] for media, cs in media_creatives.items())
        note = (
            "'신규 소재 성과' 탭의 매체 필터를 META/Tiktok/Moloco로 순회하며 읽었습니다. "
            "UAC·네이버BSA는 이 탭에 소재 단위 추적이 없어 필터 옵션에도 없는데, "
            "대신 3번 매체별 효율(MediaMix)의 매체 믹스 수치로 보충했습니다."
        )
    else:
        note = "매체별 소재 효율을 읽지 못했습니다 — 필터 셀 접근에 실패했을 수 있습니다."

    for keyword in ("UAC", "BSA"):
        m = find_media(mediamix, keyword)
        if m:
            summary = (
                f"소재 단위 추적 없음 · 매체 믹스 기준 예산 {fmt_num(m['예산'], '원')} · "
                f"예상 CPI {fmt_num(m['CPI'], '원')} · 예상 CPA {fmt_num(m['CPA'], '원')}"
            )
            rows.append([m["매체"], summary])

    children.append(para(note))
    if rows:
        children.append(table(["매체", "소재별 효율"], rows))

    children.append(h2("6. 💡 개선점"))
    improvements = []
    if low:
        names = ", ".join(c["name"] for c in low)
        improvements.append(f"🗑️ 저효율 소재 정리 검토 대상: {names}")
    if high:
        top = high[0]
        improvements.append(f"🚀 고효율 패턴 확장 후보: {top['name']} (CPA {fmt_num(top['cpa'], '원')}, USP: {top['usp']})")
    improvements.append("🔌 매체별 실적 데이터가 없어 채널 간 예산 재배분 판단이 어렵습니다 — 매체 API 연동을 우선 검토하세요.")
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


def notion_upload_file(path, content_type="image/png"):
    """노션의 파일 직접 업로드 API로 로컬 파일(차트 PNG 등)을 올리고
    file_upload id를 돌려준다. 이 id를 image 블록에서 참조하면 된다."""
    headers = {
        "Authorization": f"Bearer {os.environ['NOTION_TOKEN']}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }
    resp = requests.post(f"{NOTION_API}/file_uploads", headers=headers, json={}, timeout=15)
    resp.raise_for_status()
    upload = resp.json()

    with open(path, "rb") as f:
        send_headers = {
            "Authorization": f"Bearer {os.environ['NOTION_TOKEN']}",
            "Notion-Version": NOTION_VERSION,
        }
        r2 = requests.post(
            upload["upload_url"], headers=send_headers,
            files={"file": (os.path.basename(path), f, content_type)}, timeout=60,
        )
    r2.raise_for_status()
    return upload["id"]


def image_upload_block(file_upload_id, caption=""):
    block = {"object": "block", "type": "image", "image": {"type": "file_upload", "file_upload": {"id": file_upload_id}}}
    if caption:
        block["image"]["caption"] = [{"type": "text", "text": {"content": caption}}]
    return block


def fetch_published_dates():
    """일간 리포트 DB에 이미 만들어진 페이지들의 '날짜' 제목을 모두 모은다.
    자동 모드에서 이미 발행된 날짜를 다시 만들지 않기 위해 쓴다."""
    headers = {
        "Authorization": f"Bearer {os.environ['NOTION_TOKEN']}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }
    dates = set()
    cursor = None
    while True:
        payload = {"page_size": 100}
        if cursor:
            payload["start_cursor"] = cursor
        resp = requests.post(
            f"{NOTION_API}/databases/{DAILY_DB_ID}/query", headers=headers, json=payload, timeout=30
        )
        resp.raise_for_status()
        data = resp.json()
        for page in data.get("results", []):
            title = page.get("properties", {}).get("날짜", {}).get("title", [])
            text = "".join(t.get("plain_text", "") for t in title)
            try:
                dates.add(datetime.date.fromisoformat(text))
            except ValueError:
                continue
        if not data.get("has_more"):
            break
        cursor = data.get("next_cursor")
    return dates


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


def generate_report(sheets, drive, target_date):
    daily = fetch_daily_rows(sheets, target_date)
    mtd = fetch_mtd_summary(sheets, target_date)

    mtd_chart_id = None
    try:
        series = fetch_mtd_series(sheets, target_date)
        if series:
            chart_path = f"/tmp/mtd_chart_{target_date.isoformat()}.png"
            charts.line_chart(
                [s["date"] for s in series],
                {"앱설치": [s["install"] or 0 for s in series], "회원가입": [s["signup"] or 0 for s in series]},
                f"{target_date.year}-{target_date.month:02d} 일별 앱설치 · 회원가입 추이",
                chart_path,
            )
            mtd_chart_id = notion_upload_file(chart_path)
    except Exception as exc:  # noqa: BLE001
        print(f"[경고] MTD 차트 생성/업로드 실패 ({exc}) — 차트 없이 진행합니다.")

    age_chart_id = None
    try:
        by_age_gender = fetch_meta_signup_age_gender(target_date)
        if any(any(g.values()) for g in by_age_gender.values()):
            chart_path = f"/tmp/meta_age_chart_{target_date.isoformat()}.png"
            charts.stacked_bar_chart(
                AGE_BUCKET_ORDER,
                {
                    GENDER_LABEL_KR[gender]: [by_age_gender[age][gender] for age in AGE_BUCKET_ORDER]
                    for gender in ("female", "male", "unknown")
                },
                f"META 가입자 연령대·성별 분포 (이번 달 누적, {target_date.year}-{target_date.month:02d})",
                chart_path,
                ylabel="회원가입 수",
            )
            age_chart_id = notion_upload_file(chart_path)
    except Exception as exc:  # noqa: BLE001
        print(f"[경고] META 연령대 차트 생성/업로드 실패 ({exc}) — 차트 없이 진행합니다.")

    try:
        airbridge_actuals = fetch_airbridge_channel_actuals(target_date, target_date)
    except Exception as exc:  # noqa: BLE001
        print(f"[경고] Airbridge 채널별 실적 조회 실패 ({exc}) — 이 표는 비워둡니다.")
        airbridge_actuals = None

    media_chart_id = None
    try:
        month_start = target_date.replace(day=1)
        mtd_actuals = (
            airbridge_actuals
            if month_start == target_date
            else fetch_airbridge_channel_actuals(month_start, target_date)
        )
        if mtd_actuals:
            sorted_items = sorted(mtd_actuals.items(), key=lambda kv: kv[1]["install"] or 0, reverse=True)
            chart_path = f"/tmp/media_install_chart_{target_date.isoformat()}.png"
            charts.bar_chart(
                [label for label, _ in sorted_items],
                [v["install"] or 0 for _, v in sorted_items],
                f"매체별 실제 설치 수 (이번 달 누적, {target_date.year}-{target_date.month:02d})",
                chart_path,
                ylabel="설치 수",
            )
            media_chart_id = notion_upload_file(chart_path)
    except Exception as exc:  # noqa: BLE001
        print(f"[경고] 매체별 설치 차트 생성/업로드 실패 ({exc}) — 차트 없이 진행합니다.")

    mediamix = aggregate_mediamix(fetch_mediamix(sheets, target_date))
    active = fetch_active_creatives(sheets, target_date)
    drive_ok = True
    for c in active:
        if not drive_ok:
            c["link"] = None
            c["thumbnail"] = None
            continue
        try:
            info = drive_file_info(drive, c["name"])
            c["link"] = info["link"]
            c["thumbnail"] = info["thumbnail"]
        except Exception as exc:  # noqa: BLE001
            print(f"[경고] 드라이브 조회 실패 ({exc}) — 소재 링크는 비워둡니다.")
            drive_ok = False
            c["link"] = None
            c["thumbnail"] = None
    high, low = rank_creatives(active)

    try:
        media_creatives = fetch_creatives_by_media(sheets, target_date)
    except Exception as exc:  # noqa: BLE001
        print(f"[경고] 매체별 소재 필터 순회 실패 ({exc}) — 5번 섹션은 비워둡니다.")
        media_creatives = {}

    properties = build_properties(target_date, daily)
    children = build_children(
        target_date, daily, mediamix, high, low, media_creatives, mtd,
        mtd_chart_id, age_chart_id, airbridge_actuals, media_chart_id,
    )

    page = create_notion_page(properties, children)
    page_url = page.get("url", "")
    print(f"{target_date.isoformat()} 노션 페이지 생성 완료: {page_url}")
    notify_gchat(page_url, target_date)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="이 날짜 하나만 발행 (백필/테스트용). 생략하면 자동 모드로 동작.")
    args = parser.parse_args()

    creds = google_creds()
    sheets = sheets_client(creds)
    drive = drive_client(creds)

    if args.date:
        generate_report(sheets, drive, datetime.date.fromisoformat(args.date))
        return

    now = datetime.datetime.now(KST)
    published = fetch_published_dates()
    dates = pending_dates(now, published)
    if not dates:
        print(f"자동 모드: {now.isoformat()} 기준 새로 발행할 날짜가 없습니다.")
        return

    print(f"자동 모드: 발행 대상 {len(dates)}일 — {[d.isoformat() for d in dates]}")
    failed = []
    for d in dates:
        try:
            generate_report(sheets, drive, d)
        except Exception as exc:  # noqa: BLE001
            print(f"[실패] {d.isoformat()}: {exc}", file=sys.stderr)
            failed.append(d)
    if failed:
        raise RuntimeError(f"{len(failed)}일 발행 실패: {[d.isoformat() for d in failed]}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"실패: {exc}", file=sys.stderr)
        sys.exit(1)

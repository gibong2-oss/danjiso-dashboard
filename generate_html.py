# -*- coding: utf-8 -*-
"""
단지서비스팀 대시보드 — Google Sheets (전체 탭) → v{n+1}.html 자동 생성

탭 → D 객체 매핑:
  전자투표        → vote 시계열 + svc_usage.vote_cnt/refs
  설문조사        → survey 시계열 + svc_usage.survey_cnt/refs
  알림서비스      → notify 시계열 + svc_usage.notify_cnt
  무료단지        → 매출 제외 코드셋
  직방전환단지    → vote 20원/건 코드셋
  프로모션단지    → 프로모션 비율 맵
  정산현황        → revenue_settled_values/types
  단지별 MAU      → D.mau_m, D.top20_mau
  Rawdata         → D.table, D.autocomplete, D.svc_usage
  사업자정보      → D.table 이메일/연락처 보완
  단지 정보       → Rawdata 없을 때 폴백

  기존 HTML 유지(앱 데이터):
    D.gr, D.activity_dist_*, D.activity_monthly_*, D.complex_activity_top200

실행: python generate_html.py
"""

import re, json, gspread, os, glob, sys
from collections import defaultdict
from datetime import datetime
from google.oauth2.service_account import Credentials

# ── 설정 ─────────────────────────────────────────────────────
BASE_DIR         = os.path.dirname(os.path.abspath(__file__))
SPREADSHEET_ID   = "1Aq5wxM4J8eW2zo9_euCI5odRlb4FFJMQMbL0vWnwM4I"
CREDENTIALS_FILE = os.path.join(BASE_DIR, "service_account.json")
SCOPES           = ['https://www.googleapis.com/auth/spreadsheets']
HTML_PREFIX      = "단지서비스팀_대시보드_v"

# ── 단가 / 계산 상수 ─────────────────────────────────────────
VOTE_PRICE         = 40
VOTE_JIKBANG_PRICE = 20
SURVEY_PRICE       = 40
NOTIFY_PRICES = {
    '카카오': 20, '카카오톡': 20, 'kakao': 20, 'kakaotalk': 20,
    'sms': 20, 'lms': 40, 'mms': 80,
}
NOTIFY_DEFAULT_PRICE = 20
VAT          = 1.1
SETTLE_START = '2026-03'
DISCOUNT_30_START = '2026-05'
DISCOUNT_30_END   = '2026-07'
DISCOUNT_30_TAB   = '30% 할인 적용 단지'
SOURCE_VERSION    = 113

# 데모 / 내부 테스트 단지 — 모든 집계에서 제외
EXCLUDED_DEMO_CODES = {
    'V22222225',  # 두꺼비세상 데모
    'V16003123',  # 아파트너 (내부)
    'V22222221',  # aptner v2dev
}

# 수동 무료 처리 목록 (특정 단지 × 특정 월 → 전자투표 청구 0원)
# 매칭 조건: ym + (code 또는 biz 번호 또는 name) 중 하나 이상
MANUAL_FREE_VOTE = [
    {'ym': '2026-06', 'code': 'A10024163', 'biz': '1228287735', 'name': '검단신도시푸르지오더베뉴'},
]

def is_manual_free_vote(ym, code, name=''):
    for f in MANUAL_FREE_VOTE:
        if f['ym'] != ym: continue
        if f.get('code') == code or f.get('biz') == code or (f.get('name') and name and f['name'] == name):
            return True
    return False


# ── 활성도 기준 (MAU / 세대수 %) ─────────────────────────────
def classify_status(ratio):
    if ratio is None: return '미이용'
    if ratio >= 100:  return '매우 활성화'
    if ratio >= 70:   return '활성화'
    if ratio >= 40:   return '사용'
    if ratio >= 10:   return '비활성'
    if ratio > 0:     return '이탈'
    return '미이용'


# ═══════════════════════════════════════════════════════════════
# 1. Sheets 연결 & 탭 읽기 헬퍼
# ═══════════════════════════════════════════════════════════════
def connect_sheets():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


def read_tab(ss, tab_name):
    """탭 전체 읽기 → (header_list, rows_list)"""
    try:
        ws   = ss.worksheet(tab_name)
        vals = ws.get_all_values()
        if not vals:
            return [], []
        return vals[0], vals[1:]
    except Exception as e:
        print(f"  ⚠ [{tab_name}] 읽기 실패: {e}")
        return [], []


def to_dicts(header, rows):
    """header + rows → list[dict]"""
    result = []
    for row in rows:
        if not any(str(v).strip() for v in row):
            continue
        d = {header[i]: (row[i] if i < len(row) else '') for i in range(len(header))}
        result.append(d)
    return result


def safe_int(v, default=0):
    try:
        return int(float(str(v).replace(',', '').strip()))
    except:
        return default


def safe_float(v, default=None):
    try:
        return round(float(str(v).replace(',', '').strip()), 2)
    except:
        return default


def ym_of(date_str):
    """'YYYY-MM-DD ...' → 'YYYY-MM'  (None if invalid)"""
    s = str(date_str).strip()[:7]
    return s if re.match(r'\d{4}-\d{2}$', s) else None


def find_code_col(header):
    for h in header:
        if '코드' in h or h.lower() in ('kapt_code', 'code'):
            return h
    return header[0] if header else None


# ═══════════════════════════════════════════════════════════════
# 2. 참조 탭 로더
# ═══════════════════════════════════════════════════════════════
def load_code_set(ss, tab_name):
    """단지코드 집합 (무료단지 등)"""
    header, rows = read_tab(ss, tab_name)
    codes = set()
    if not header:
        return codes
    col = find_code_col(header)
    if not col:
        return codes
    idx = header.index(col)
    for row in rows:
        c = row[idx].strip() if idx < len(row) else ''
        if c:
            codes.add(c)
    return codes


def load_jikbang(ss):
    """직방전환단지: set of (code, ym) — 기준월 기준으로 해당 월에만 20원/건 적용.
    단, SETTLE_START(2026-03) 이전 월은 제외 (정산 이전 기간은 전원 40원 적용).
    """
    header, rows = read_tab(ss, '직방전환단지')
    pairs = set()
    if not header:
        return pairs
    code_col = find_code_col(header)
    ym_col   = next((h for h in header if '기준월' in str(h)), None)
    if not code_col or not ym_col:
        return pairs
    ci = header.index(code_col)
    yi = header.index(ym_col)
    for row in rows:
        code = row[ci].strip() if ci < len(row) else ''
        ym   = str(row[yi]).strip()[:7] if yi < len(row) else ''  # YYYY-MM
        if code and re.match(r'\d{4}-\d{2}', ym) and ym >= SETTLE_START:
            pairs.add((code, ym))
    return pairs


def load_promo(ss):
    """Promotion sheet (kept per-month snapshot).
    Columns: A=base_month, B=code, C=name, D=remainQty, E=memo, F=startDate
    Returns: {(code, 'YYYY-MM'): {'remain': int, 'startDate': 'YYYY-MM-DD' or ''}}
    """
    header, rows = read_tab(ss, '프로모션단지')
    promo = {}
    if not header:
        return promo
    # column finders
    def find_col(keys):
        for h in header:
            if any(k in h for k in keys):
                return header.index(h)
        return -1
    code_idx  = header.index(find_code_col(header)) if find_code_col(header) in header else -1
    ym_idx    = find_col(['기준월', 'base'])
    remain_idx= find_col(['잔여수량', 'remain'])
    start_idx = find_col(['적용시작일', '시작일', 'startDate', 'start'])
    if code_idx < 0 or ym_idx < 0 or remain_idx < 0:
        return promo
    for row in rows:
        if not any(str(v).strip() for v in row):
            continue
        code = str(row[code_idx]).strip() if code_idx < len(row) else ''
        ym_raw = str(row[ym_idx]).strip()[:7] if ym_idx < len(row) else ''
        if not code or not re.match(r'\d{4}-\d{2}', ym_raw):
            continue
        try:
            remain = int(float(str(row[remain_idx]).replace(',', '').strip())) if remain_idx < len(row) else 0
        except:
            remain = 0
        start = str(row[start_idx]).strip()[:10] if start_idx >= 0 and start_idx < len(row) else ''
        promo[(code, ym_raw)] = {'remain': remain, 'startDate': start}
    return promo


def load_promo_legacy(ss):
    """Backward-compat: {code: 1.0} for aggregate_* functions (no rate column in new sheet)."""
    p = load_promo(ss)
    return {code: 1.0 for (code, _ym) in p.keys()}


def load_discount_30(ss):
    """30% 할인 적용 단지 코드 set."""
    header, rows = read_tab(ss, DISCOUNT_30_TAB)
    codes = set()
    if not header:
        return codes
    col = find_code_col(header)
    if not col:
        return codes
    idx = header.index(col)
    for row in rows:
        c = row[idx].strip() if idx < len(row) else ''
        if c:
            codes.add(c)
    return codes


def load_prepaid_promo(ss):
    """선결제 프로모션 구매 현황 시트.
    헤더로 컬럼 자동 매칭 (구조 변경 대응):
      - 날짜: '처리 완료일' / '실 적용 시작 일' / '구매 완료일' / 'A열 (fallback)'
      - 매출: '결제 금액의 SUM' / '결제 금액' / '부가세 포함 매출' / 'E열 (fallback)'
    Returns: {months, values, weeks, wk_values}  (만원 단위, VAT 포함)
    """
    header, rows = read_tab(ss, '선결제 프로모션 구매 현황')
    if not header:
        return {'months': [], 'values': [], 'weeks': [], 'wk_values': []}
    # 컬럼 자동 매칭
    def _find_col(keys):
        for k in keys:
            if k in header: return header.index(k)
        for i, h in enumerate(header):
            for k in keys:
                if k in h: return i
        return -1
    date_idx = _find_col(['처리 완료일', '구매 완료일', '구매완료일', '실 적용 시작 일', '실적용시작일', '발행일'])
    amt_idx = _find_col(['결제 금액의 SUM', '결제금액의 SUM', '결제 금액', '결제금액', '부가세 포함 매출'])
    if date_idx < 0: date_idx = 0   # fallback (A열)
    if amt_idx < 0: amt_idx = 4    # fallback (E열)
    print(f"  선결제 시트 컬럼: 날짜={header[date_idx] if date_idx<len(header) else '?'} (idx={date_idx}), 매출={header[amt_idx] if amt_idx<len(header) else '?'} (idx={amt_idx})")
    by_month = defaultdict(float)
    by_week = defaultdict(float)
    by_day = defaultdict(float)
    cnt_month = defaultdict(int)
    cnt_week = defaultdict(int)
    skipped_date = 0
    skipped_amt = 0
    skipped_len = 0
    for row in rows:
        if len(row) <= max(date_idx, amt_idx):
            skipped_len += 1
            continue
        raw_date = str(row[date_idx]).strip()
        norm_date = re.sub(r'[./\s]+', '-', raw_date)
        m_date = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', norm_date)
        if not m_date:
            skipped_date += 1
            continue
        date_str = f"{m_date.group(1)}-{int(m_date.group(2)):02d}-{int(m_date.group(3)):02d}"
        ym = date_str[:7]
        wk = _date_to_isoweek(date_str)
        amt_raw = str(row[amt_idx]).replace(',', '').replace('원', '').replace('₩', '').strip()
        try:
            amt = float(amt_raw or 0)
        except:
            skipped_amt += 1
            amt = 0
        amt_man = amt / 10000
        by_month[ym] += amt_man
        by_week[wk] += amt_man
        by_day[date_str] += amt_man
        cnt_month[ym] += 1
        cnt_week[wk] += 1
    print(f"  선결제 프로모션: 총 {len(rows)}행, 유효 {sum(cnt_month.values())}건, 스킵(짧음={skipped_len}, 날짜={skipped_date}, 금액={skipped_amt})")
    months = sorted(by_month.keys())
    weeks = sorted(by_week.keys())
    days = sorted(by_day.keys())
    return {
        'months': months,
        'values': [round(by_month[m], 2) for m in months],
        'weeks': weeks,
        'wk_values': [round(by_week[w], 2) for w in weeks],
        'days': days,
        'day_values': [round(by_day[d], 2) for d in days],
        'monthly_count': {m: cnt_month[m] for m in months},
        'weekly_count': {w: cnt_week[w] for w in weeks}
    }


def load_officener_rev(ss):
    """오피스너 매출 — 발행일/금액(VAT별도)/수익모델 컬럼.
    화면 표시는 VAT 포함 (× 1.1).
    Returns: {months, values, by_type, by_month_type, by_week, by_week_type}
    """
    header, rows = read_tab(ss, '오피스너')
    if not header:
        return {'months': [], 'values': [], 'by_type': {}, 'by_month_type': {}, 'weeks': [], 'wk_values': [], 'by_week_type': {}}
    
    # 컬럼 찾기 — exact match 우선, 그 다음 partial
    def find_col(keys):
        for k in keys:
            if k in header: return header.index(k)
        for h in header:
            for k in keys:
                if k in h: return header.index(h)
        return -1
    date_idx = find_col(['발행일'])
    amt_idx  = find_col(['금액(VAT별도)', '금액'])
    type_idx = find_col(['수익모델', '모델'])
    status_idx = find_col(['발행 상태', '발행상태'])  # 발행 완료 필터용
    if date_idx < 0 or amt_idx < 0:
        return {'months': [], 'values': [], 'by_type': {}, 'by_month_type': {}, 'weeks': [], 'wk_values': [], 'by_week_type': {}}
    
    by_month = defaultdict(float)         # ym → amt (만원, VAT 포함)
    by_month_type = defaultdict(lambda: defaultdict(float))  # ym → type → amt
    by_week = defaultdict(float)          # YYYY-Www → amt
    by_week_type = defaultdict(lambda: defaultdict(float))
    by_type_total = defaultdict(float)
    
    for row in rows:
        if date_idx >= len(row) or amt_idx >= len(row): continue
        # 발행 상태 필터 — '발행 완료'만 매출 인정
        if status_idx >= 0 and status_idx < len(row):
            stat = str(row[status_idx]).strip()
            if stat not in ('발행완료', '발행 완료'): continue
        date_str = str(row[date_idx]).strip()[:10]
        if not re.match(r'\d{4}-\d{2}-\d{2}', date_str): continue
        ym = date_str[:7]
        wk = _date_to_isoweek(date_str)
        try:
            amt_nv = float(str(row[amt_idx]).replace(',', '').replace('원', '').strip() or 0)
        except:
            continue
        amt = amt_nv * VAT  # VAT 포함 → 원 단위
        amt_man = amt / 10000  # 만원
        typ = str(row[type_idx]).strip() if type_idx >= 0 and type_idx < len(row) else ''
        typ = typ or '기타'
        by_month[ym] += amt_man
        by_month_type[ym][typ] += amt_man
        by_week[wk] += amt_man
        by_week_type[wk][typ] += amt_man
        by_type_total[typ] += amt_man
    
    months = sorted(by_month.keys())
    weeks = sorted(by_week.keys())
    return {
        'months': months,
        'values': [round(by_month[m], 2) for m in months],
        'weeks': weeks,
        'wk_values': [round(by_week[w], 2) for w in weeks],
        'by_type': {t: round(by_type_total[t], 2) for t in by_type_total},
        'by_month_type': {ym: {t: round(by_month_type[ym][t], 2) for t in by_month_type[ym]} for ym in by_month_type},
        'by_week_type': {wk: {t: round(by_week_type[wk][t], 2) for t in by_week_type[wk]} for wk in by_week_type}
    }


def load_aptstory_rev(ss):
    """아파트스토리 매출 — 매출구분=매출인 행만. 합계금액/품목명 컬럼.
    합계금액 = 매출 (원, 그대로). 화면 표시 시 그대로.
    """
    header, rows = read_tab(ss, '아파트스토리')
    if not header:
        return {'months': [], 'values': [], 'by_type': {}, 'by_month_type': {}, 'weeks': [], 'wk_values': [], 'by_week_type': {}}
    
    def find_col(keys):
        for k in keys:
            if k in header: return header.index(k)
        for h in header:
            for k in keys:
                if k in h: return header.index(h)
        return -1
    date_idx = find_col(['작성일자'])
    kind_idx = find_col(['매출구분'])
    amt_idx  = find_col(['합계금액'])
    type_idx = find_col(['구분'])  # AL열 구분 컬럼
    if date_idx < 0 or amt_idx < 0 or kind_idx < 0:
        return {'months': [], 'values': [], 'by_type': {}, 'by_month_type': {}, 'weeks': [], 'wk_values': [], 'by_week_type': {}}
    
    by_month = defaultdict(float)
    by_month_type = defaultdict(lambda: defaultdict(float))
    by_week = defaultdict(float)
    by_week_type = defaultdict(lambda: defaultdict(float))
    by_type_total = defaultdict(float)
    
    for row in rows:
        if date_idx >= len(row) or amt_idx >= len(row) or kind_idx >= len(row): continue
        kind = str(row[kind_idx]).strip()
        if kind != '매출': continue  # 매출구분 = '매출' 행만 (시트 정리 완료)
        date_str = str(row[date_idx]).strip()[:10]
        if not re.match(r'\d{4}-\d{2}-\d{2}', date_str): continue
        ym = date_str[:7]
        wk = _date_to_isoweek(date_str)
        try:
            amt = float(str(row[amt_idx]).replace(',', '').replace('원', '').strip() or 0)
        except:
            continue
        amt_man = amt / 10000  # 만원
        typ = str(row[type_idx]).strip() if type_idx >= 0 and type_idx < len(row) else ''
        # 4종 카테고리만 매출 인정 (문자충전/앱이용료/이사예약/전자투표)
        ALLOWED_AS_KINDS = {'문자충전', '앱이용료', '이사예약', '전자투표'}
        if typ not in ALLOWED_AS_KINDS:
            continue
        by_month[ym] += amt_man
        by_month_type[ym][typ] += amt_man
        by_week[wk] += amt_man
        by_week_type[wk][typ] += amt_man
        by_type_total[typ] += amt_man
    
    months = sorted(by_month.keys())
    weeks = sorted(by_week.keys())
    return {
        'months': months,
        'values': [round(by_month[m], 2) for m in months],
        'weeks': weeks,
        'wk_values': [round(by_week[w], 2) for w in weeks],
        'by_type': {t: round(by_type_total[t], 2) for t in by_type_total},
        'by_month_type': {ym: {t: round(by_month_type[ym][t], 2) for t in by_month_type[ym]} for ym in by_month_type},
        'by_week_type': {wk: {t: round(by_week_type[wk][t], 2) for t in by_week_type[wk]} for wk in by_week_type}
    }


def load_settlement(ss):
    """{svc: {YYYY-MM: 만원(VAT포함)}} — 서비스별 월 청구금액 합산
    컬럼 구조: 정산월, 단지코드, 단지명, 서비스명, 구분, 프로모션참여,
               사용건수, 무료처리, 실청구건수, 단가, 공급가액, 부가세, 청구금액, 발행여부
    - 청구금액: 원 단위 → ÷10000 하여 만원으로 저장
    - 발행여부 == 'ok' 인 행만 집계
    - 서비스별로 (svc, ym) 단위로 SUM
    """
    header, rows = read_tab(ss, '정산현황')
    SVC_MAP = {'전자투표': 'vote', '전자투표 (30% 할인)': 'vote', '알림서비스': 'notify', '설문조사': 'survey'}
    settled = {'vote': {}, 'notify': {}, 'survey': {}}
    if not header:
        return settled
    for d in to_dicts(header, rows):
        # 발행여부 필터 완화 (cancel/void/취소/무효/skip 제외)
        issued = str(d.get('발행여부', '')).strip().lower()
        if issued in ('cancel', 'void', '취소', '무효', 'skip'):
            continue
        # 정산월 파싱
        ym_raw = str(d.get('정산월', '')).strip()
        m = re.match(r'(\d{4})[./-](\d{1,2})', ym_raw)
        if not m:
            continue
        ym_key = f"{m.group(1)}-{int(m.group(2)):02d}"
        # 단지코드 (수동 무료 처리용)
        code = str(d.get('단지코드', '')).strip()
        # 서비스명 매핑
        svc = SVC_MAP.get(str(d.get('서비스명', '')).strip())
        if not svc:
            continue
        # 청구금액 (원 단위) → 만원 — 수동 무료 처리 시 0
        try:
            amt = float(str(d.get('청구금액', 0)).replace(',', '').replace('원', '').strip())
            # 전자투표 + 수동 무료 처리 시 0
            if svc == 'vote' and is_manual_free_vote(ym_key, code):
                amt = 0
            settled[svc][ym_key] = round(settled[svc].get(ym_key, 0) + amt / 10000, 2)
        except Exception:
            pass
    return settled


def load_settlement_from_raw(ss, free, jikbang, promo_snap, discount_30_codes):
    """V19 buildBillingData ported to Python (외부 자동정산 대시보드와 일치).
    핵심: 단지별 vote 사용량 잔여수량 carryover, 발송일자별 30% 할인 row 단위 적용,
          단지별 service-group VAT 절사, MIN_INVOICE 1000원 미만 skip.
    """
    MIN_INV = 1000
    DISC_S = '2026-05-01'; DISC_E = '2026-07-31'
    EXCLUDED = [{'code': 'C41830251', 's': '2026-05-07', 'e': '2026-05-08'}]
    def is_excluded(code, dstr):
        for r in EXCLUDED:
            if r['code'] == code and r['s'] <= dstr <= r['e']:
                return True
        return False

    # ----- 1. raw rows load -----
    all_vote_rows = []  # full list for carryover
    header, rows = read_tab(ss, '전자투표')
    for d in to_dicts(header, rows):
        date_str = str(d.get('일자', '')).strip()[:10]
        ym = ym_of(date_str)
        code = str(d.get('단지코드', '')).strip()
        if code in EXCLUDED_DEMO_CODES: continue
        c = safe_int(d.get('건수', 0))
        if not ym or not code or c <= 0: continue
        if is_excluded(code, date_str): continue
        all_vote_rows.append({'date': date_str, 'ym': ym, 'code': code, 'count': c})

    notify_rows = []
    header, rows = read_tab(ss, '알림서비스')
    for d in to_dicts(header, rows):
        date_str = str(d.get('일자', '')).strip()[:10]
        ym = ym_of(date_str)
        code = str(d.get('단지코드', '')).strip()
        if code in EXCLUDED_DEMO_CODES: continue
        c = safe_int(d.get('건수', 0))
        stype = str(d.get('발송수단', '')).strip().lower()
        if not ym or not code or c <= 0: continue
        if ym < SETTLE_START: continue
        notify_rows.append({'ym': ym, 'code': code, 'count': c, 'rate': NOTIFY_PRICES.get(stype, NOTIFY_DEFAULT_PRICE)})

    survey_rows = []
    header, rows = read_tab(ss, '설문조사')
    for d in to_dicts(header, rows):
        date_str = str(d.get('일자', '')).strip()[:10]
        ym = ym_of(date_str)
        code = str(d.get('단지코드', '')).strip()
        if code in EXCLUDED_DEMO_CODES: continue
        c = safe_int(d.get('건수', 0))
        if not ym or not code or c <= 0: continue
        if ym < SETTLE_START: continue
        survey_rows.append({'ym': ym, 'code': code, 'count': c})

    # ----- 2. carry-over baseline: per-code earliest promo_snap row -----
    promo_baseline = {}  # {code: (ym, remainQty, startDate)}
    for (code, ym), entry in promo_snap.items():
        cur = promo_baseline.get(code)
        if cur is None or ym < cur[0]:
            promo_baseline[code] = (ym, entry['remain'], entry.get('startDate', ''))

    def compute_carryover(code, target_ym):
        """V19 v17ComputePromoCarryover ported."""
        b = promo_baseline.get(code)
        if b is None:
            return None
        base_month, baseline_qty, base_start = b
        if base_month > target_ym:
            return None
        target_first = target_ym + '-01'
        total_usage = 0
        if base_start:
            for vr in all_vote_rows:
                if vr['code'] != code: continue
                if vr['date'] >= base_start and vr['date'] < target_first:
                    total_usage += vr['count']
        else:
            for vr in all_vote_rows:
                if vr['code'] != code: continue
                if vr['ym'] >= base_month and vr['ym'] < target_ym:
                    total_usage += vr['count']
        return baseline_qty - total_usage

    # ----- 3. group rows by (ym, code) -----
    vote_by_cm = defaultdict(list)
    for vr in all_vote_rows:
        if vr['ym'] >= SETTLE_START:
            vote_by_cm[(vr['ym'], vr['code'])].append(vr)
    notify_by_cm = defaultdict(list)
    for r in notify_rows:
        notify_by_cm[(r['ym'], r['code'])].append(r)
    survey_by_cm = defaultdict(list)
    for r in survey_rows:
        survey_by_cm[(r['ym'], r['code'])].append(r)

    all_cms = set(vote_by_cm.keys()) | set(notify_by_cm.keys()) | set(survey_by_cm.keys())

    # ----- 4. bill each (ym, code) -----
    settled = {'vote': defaultdict(float), 'notify': defaultdict(float), 'survey': defaultdict(float)}

    for (ym, code) in all_cms:
        is_free = code in free
        promo_entry = promo_baseline.get(code)
        is_promo = promo_entry is not None
        is_jikbang = (code, ym) in jikbang
        effective_vote_rate = 0 if is_free else (VOTE_JIKBANG_PRICE if is_jikbang else VOTE_PRICE)

        # vote rows of this (ym, code)
        rows_v = vote_by_cm.get((ym, code), [])
        promo_start = promo_entry[2] if is_promo else ''
        eligible = [r for r in rows_v if not promo_start or r['date'] >= promo_start]
        pre_rows = [r for r in rows_v if promo_start and r['date'] < promo_start]
        total_use = sum(r['count'] for r in eligible)
        pre_use = sum(r['count'] for r in pre_rows)

        # billed_vote
        if is_free:
            billed_vote = 0
        elif is_promo and total_use > 0:
            snap = compute_carryover(code, ym)
            snap = snap if snap is not None else 0
            if snap <= 0:
                billed_vote = total_use + pre_use
            else:
                after = snap - total_use
                billed_vote = (abs(after) if after < 0 else 0) + pre_use
        elif is_promo and total_use == 0 and pre_use > 0:
            billed_vote = pre_use
        else:
            billed_vote = total_use + pre_use

        # svcMap: group by (rate, dc_flag)
        # dc_flag = row in 30% discount range
        is_dc_code = code in discount_30_codes
        svc_groups = {}  # key: (rate, dc_flag) -> use_count
        for r in rows_v:
            dc = is_dc_code and (DISC_S <= r['date'] <= DISC_E)
            key = (effective_vote_rate, dc)
            svc_groups[key] = svc_groups.get(key, 0) + r['count']

        # vote_supply: 비례 분배 + dc_mul
        all_use = total_use + pre_use
        vote_supply = 0.0
        if all_use > 0:
            entries = list(svc_groups.items())
            billed_allocated = 0
            for vi, ((rate, dc), use_count) in enumerate(entries):
                if is_promo or is_free:
                    if vi == len(entries) - 1:
                        row_billed = billed_vote - billed_allocated
                    else:
                        row_billed = round(use_count * billed_vote / all_use) if all_use > 0 else 0
                        billed_allocated += row_billed
                else:
                    row_billed = use_count
                dc_mul = 0.7 if dc else 1.0
                vote_supply += max(0, row_billed) * rate * dc_mul

        # notify_supply, survey_supply
        notify_supply = sum(r['count'] * r['rate'] for r in notify_by_cm.get((ym, code), []))
        survey_supply = sum(r['count'] * SURVEY_PRICE for r in survey_by_cm.get((ym, code), []))

        # VAT floor + MIN_INVOICE skip
        def floor_vat(s):
            return int(s * 0.1 / 10) * 10
        for svc, sup in [('vote', vote_supply), ('notify', notify_supply), ('survey', survey_supply)]:
            if sup <= 0: continue
            vat = floor_vat(sup)
            total = sup + vat
            if total < MIN_INV: continue
            settled[svc][ym] += total / 10000

    return {svc: {ym: round(v, 2) for ym, v in dd.items()} for svc, dd in settled.items()}


def aggregate_vote(ss, free, jikbang, promo):
    """전자투표 집계.
    무료단지(free)는 완전 제외 — 미발행에도 포함 안 함.
    유료단지 중 단지별 월 청구금액 < 1,000원(=0.1만원)인 경우 미발행으로 분류.
    2-pass: 1st) (ym, code)별 합산, 2nd) threshold 분기
    """
    THRESHOLD = 0.1          # 만원 = 1,000원
    header, rows = read_tab(ss, '전자투표')

    # ── 1st pass: (ym, code) 단위 집계 (무료단지 포함, revenue=0으로 처리) ──
    cplx_mo = defaultdict(lambda: {'name': '', 'cnt': 0, 'rev': 0.0, 'rows': [], 'is_free': False})
    for d in to_dicts(header, rows):
        date_str = str(d.get('일자', '')).strip()[:10]
        ym   = ym_of(date_str)
        code = str(d.get('단지코드', '')).strip()
        if code in EXCLUDED_DEMO_CODES: continue
        name = str(d.get('단지명', '')).strip()
        c    = safe_int(d.get('건수', 0))
        if not ym or not code or c <= 0:
            continue
        # 무료단지 or 수동 무료 처리 (특정 단지 × 특정 월)
        is_free = code in free or is_manual_free_vote(ym, code, name)
        if is_free:
            revenue = 0.0
        elif ym < SETTLE_START:
            # 정산 이전 기간: 직방·프로모션 할인 없이 표준가 일괄 적용
            revenue = c * VOTE_PRICE * VAT / 10000
        else:
            # 정산 이후(2026-03~): 기준월 일치하는 경우만 직방 단가, 프로모션 적용
            unit    = VOTE_JIKBANG_PRICE if (code, ym) in jikbang else VOTE_PRICE
            revenue = c * unit * promo.get(code, 1.0) * VAT / 10000
        ref = str(d.get('참조ID', '')).strip()
        key = (ym, code)
        if name:
            cplx_mo[key]['name'] = name
        cplx_mo[key]['is_free'] = is_free
        cplx_mo[key]['cnt'] += c
        cplx_mo[key]['rev'] += revenue
        cplx_mo[key]['rows'].append(
            {'date': date_str, 'cnt': c, 'rev': revenue, 'ref': ref}
        )

    # ── 2nd pass: threshold 분기 ──────────────────────────────
    cnt      = defaultdict(int)
    ids      = defaultdict(set)      # 참조ID 기준 이용 횟수
    cplx_ids = defaultdict(set)      # 단지코드 기준 단지 수 (무료 포함, 단지당 매출 분모)
    rev      = defaultdict(float)
    svc_cnt  = defaultdict(int)
    svc_refs = defaultdict(set)
    day_cnt  = defaultdict(int)
    day_rev  = defaultdict(float)
    excl_rows = defaultdict(int)
    excl_sms  = defaultdict(int)
    excl_cmap = {}

    for (ym, code), data in cplx_mo.items():
        is_free = data['is_free']
        # 유료단지 중 1,000원 미만 → 미발행 추적 (무료단지는 excl 대상 아님)
        if not is_free and data['rev'] < THRESHOLD:
            excl_rows[ym] += len(data['rows'])
            excl_sms[ym]  += data['cnt']
            if ym not in excl_cmap:
                excl_cmap[ym] = {}
            excl_cmap[ym][code] = {
                'name': data['name'], 'rows': len(data['rows']), 'sms': data['cnt'],
                'rev': round(data['rev'] * 10000)
            }
        # 이용 횟수/건수는 무료단지 포함 모든 단지
        cplx_ids[ym].add(code)       # 단지 수 (무료 포함)
        for r in data['rows']:
            cnt[ym]  += r['cnt']
            if r['ref']:
                ids[ym].add(r['ref'])      # 참조ID 기준 이용 횟수
            # 추정 매출: 직방·프로모션·무료 구분 없이 표준가 일괄 적용
            rev_est = r['cnt'] * VOTE_PRICE * VAT / 10000
            rev[ym] += rev_est
            svc_cnt[code] += r['cnt']
            if r['ref']:
                svc_refs[code].add(r['ref'])
            if re.match(r'\d{4}-\d{2}-\d{2}', r['date']):
                day_cnt[r['date']] += r['cnt']
                day_rev[r['date']] += rev_est

    return cnt, ids, rev, cplx_ids, svc_cnt, svc_refs, day_cnt, day_rev, excl_rows, excl_sms, excl_cmap


def aggregate_survey(ss, free, promo):
    """설문조사 집계.
    무료단지 개념 없음 — 단지별 월 청구금액 < 1,000원(=0.1만원)인 경우 미발행으로 분류.
    2-pass: 1st) (ym, code)별 합산, 2nd) threshold 분기
    """
    THRESHOLD = 0.1          # 만원 = 1,000원
    header, rows = read_tab(ss, '설문조사')

    # ── 1st pass: (ym, code) 단위 집계 ───────────────────────
    cplx_mo = defaultdict(lambda: {'name': '', 'cnt': 0, 'rev': 0.0, 'rows': []})
    for d in to_dicts(header, rows):
        date_str = str(d.get('일자', '')).strip()[:10]
        ym   = ym_of(date_str)
        code = str(d.get('단지코드', '')).strip()
        if code in EXCLUDED_DEMO_CODES: continue
        name = str(d.get('단지명', d.get('kapt_name', d.get('apt_name', '')))).strip()
        c    = safe_int(d.get('건수', 0))
        if not ym or not code or c <= 0:
            continue
        revenue = c * SURVEY_PRICE * VAT / 10000  # 추정 매출: 프로모션 제외, 표준가
        ref = str(d.get('참조ID', '')).strip()
        key = (ym, code)
        if name:
            cplx_mo[key]['name'] = name
        cplx_mo[key]['cnt'] += c
        cplx_mo[key]['rev'] += revenue
        cplx_mo[key]['rows'].append(
            {'date': date_str, 'cnt': c, 'rev': revenue, 'ref': ref}
        )

    # ── 2nd pass: threshold 분기 ──────────────────────────────
    cnt      = defaultdict(int)
    ids      = defaultdict(set)      # 참조ID 기준 이용 횟수
    cplx_ids = defaultdict(set)      # 단지코드 기준 단지 수 (단지당 매출 분모)
    rev      = defaultdict(float)
    svc_cnt  = defaultdict(int)
    svc_refs = defaultdict(set)
    day_cnt  = defaultdict(int)
    day_rev  = defaultdict(float)
    excl_rows = defaultdict(int)
    excl_sms  = defaultdict(int)
    excl_cmap = {}

    for (ym, code), data in cplx_mo.items():
        if data['rev'] < THRESHOLD:          # 1,000원 미만 → 미발행 추적
            excl_rows[ym] += len(data['rows'])
            excl_sms[ym]  += data['cnt']
            if ym not in excl_cmap:
                excl_cmap[ym] = {}
            excl_cmap[ym][code] = {
                'name': data['name'], 'rows': len(data['rows']), 'sms': data['cnt'],
                'rev': round(data['rev'] * 10000)
            }
        # 메인 집계는 threshold 무관하게 모든 단지 포함
        cplx_ids[ym].add(code)       # 단지 수
        for r in data['rows']:
            cnt[ym]  += r['cnt']
            if r['ref']:
                ids[ym].add(r['ref'])      # 참조ID 기준 이용 횟수
            rev[ym]  += r['rev']
            svc_cnt[code] += r['cnt']
            if r['ref']:
                svc_refs[code].add(r['ref'])
            if re.match(r'\d{4}-\d{2}-\d{2}', r['date']):
                day_cnt[r['date']] += r['cnt']
                day_rev[r['date']] += r['rev']

    return cnt, ids, rev, cplx_ids, svc_cnt, svc_refs, day_cnt, day_rev, excl_rows, excl_sms, excl_cmap


def aggregate_notify(ss, free):
    """알림서비스 집계.
    무료단지 개념 없음 — 단지별 월 청구금액 < 1,000원(=0.1만원)인 경우 미발행으로 분류.
    2-pass: 1st) (ym, code)별 합산, 2nd) threshold 분기
    """
    THRESHOLD = 0.1          # 만원 = 1,000원
    header, rows = read_tab(ss, '알림서비스')

    # ── 1st pass: (ym, code) 단위 집계 ───────────────────────
    cplx_mo = defaultdict(lambda: {'name': '', 'cnt': 0, 'rev': 0.0, 'rows': []})
    for d in to_dicts(header, rows):
        date_str = str(d.get('일자', '')).strip()[:10]
        ym      = ym_of(date_str)
        code    = str(d.get('단지코드', '')).strip()
        name    = str(d.get('단지명', '')).strip()
        c       = safe_int(d.get('건수', 0))
        stype   = str(d.get('발송수단', '')).strip().lower()
        content = str(d.get('내용', '')).strip()
        if not ym or not code or c <= 0:
            continue
        unit    = NOTIFY_PRICES.get(stype, NOTIFY_DEFAULT_PRICE)
        revenue = c * unit * VAT / 10000
        key = (ym, code)
        if name:
            cplx_mo[key]['name'] = name
        cplx_mo[key]['cnt'] += c
        cplx_mo[key]['rev'] += revenue
        cplx_mo[key]['rows'].append(
            {'date': date_str, 'cnt': c, 'rev': revenue, 'content': content}
        )

    # ── 2nd pass: threshold 분기 ──────────────────────────────
    sends     = defaultdict(int)
    ids       = defaultdict(set)
    rev       = defaultdict(float)
    events    = defaultdict(set)
    svc_cnt   = defaultdict(int)
    day_cnt   = defaultdict(int)
    day_rev   = defaultdict(float)
    excl_rows = defaultdict(int)
    excl_sms  = defaultdict(int)
    excl_cmap = {}

    for (ym, code), data in cplx_mo.items():
        if data['rev'] < THRESHOLD:          # 1,000원 미만 → 미발행 추적
            excl_rows[ym] += len(data['rows'])
            excl_sms[ym]  += data['cnt']
            if ym not in excl_cmap:
                excl_cmap[ym] = {}
            excl_cmap[ym][code] = {
                'name': data['name'], 'rows': len(data['rows']), 'sms': data['cnt'],
                'rev': round(data['rev'] * 10000)
            }
        # 메인 집계는 threshold 무관하게 모든 단지 포함
        for r in data['rows']:
            sends[ym]  += r['cnt']
            ids[ym].add(code)
            rev[ym]    += r['rev']
            events[ym].add((r['date'], r['content'][:80]))
            svc_cnt[code] += r['cnt']
            if re.match(r'\d{4}-\d{2}-\d{2}', r['date']):
                day_cnt[r['date']] += r['cnt']
                day_rev[r['date']] += r['rev']

    return sends, ids, rev, events, svc_cnt, day_cnt, day_rev, excl_rows, excl_sms, excl_cmap


# ═══════════════════════════════════════════════════════════════
# 4. 단지별 MAU → D.mau_m, D.top20_mau
# ═══════════════════════════════════════════════════════════════
def load_mau(ss, all_months):
    header, rows = read_tab(ss, '단지별 MAU')
    if not header:
        return {}, []
    # {code: {name, {ym: mau}}}
    cplx = {}
    for d in to_dicts(header, rows):
        code = str(d.get('kapt_code', '')).strip()
        if code in EXCLUDED_DEMO_CODES: continue
        name = str(d.get('apt_name',  '')).strip()
        ym   = str(d.get('date', '')).strip()[:7]
        mau  = safe_int(d.get('MAU', 0))
        if not code or not ym or mau < 0:
            continue
        if code not in cplx:
            cplx[code] = {'name': name, 'monthly': {}}
        cplx[code]['monthly'][ym] = mau
        if name:
            cplx[code]['name'] = name

    # D.mau_m : {code: {name, months, vals}}
    mau_m = {}
    for code, info in cplx.items():
        ym_list = sorted(info['monthly'])
        mau_m[code] = {
            'name':   info['name'],
            'months': ym_list,
            'vals':   [info['monthly'][m] for m in ym_list],
        }

    # D.top20_mau : top 20 by lifetime MAU sum, values aligned to all_months
    scored = sorted(cplx.items(), key=lambda x: sum(x[1]['monthly'].values()), reverse=True)
    top20_mau = []
    for code, info in scored[:20]:
        data = [info['monthly'].get(m, 0) for m in all_months]
        top20_mau.append({'name': info['name'], 'data': data})

    return mau_m, top20_mau


# ═══════════════════════════════════════════════════════════════
# 5. Rawdata + 사업자정보 → D.table, D.autocomplete
# ═══════════════════════════════════════════════════════════════
def load_table_and_autocomplete(ss, mau_m):
    """
    Rawdata (또는 폴백: 단지 정보) + 사업자정보 + mau_m 으로
    D.table, D.autocomplete 생성
    """
    header, rows = read_tab(ss, 'Rawdata')
    if not header:
        print("  ⚠ Rawdata 없음 → 단지 정보 폴백")
        header, rows = read_tab(ss, '단지 정보')
    if not header:
        return [], []

    # 사업자정보 {코드: {메일주소, 업체 연락처}}
    biz_header, biz_rows = read_tab(ss, '사업자정보')
    biz_map = {}
    for bd in to_dicts(biz_header, biz_rows):
        code = str(bd.get('코드', '')).strip()
        if code:
            biz_map[code] = {
                'email': str(bd.get('메일주소', '')).strip(),
                'tel':   str(bd.get('업체 연락처', '')).strip(),
            }

    def get_latest_mau(code):
        info = mau_m.get(code, {})
        vals = info.get('vals', [])
        return vals[-1] if vals else 0

    def get_latest_ym(code):
        info = mau_m.get(code, {})
        months = info.get('months', [])
        return months[-1] if months else None

    table = []
    autocomplete = []

    for d in to_dicts(header, rows):
        code = str(d.get('kapt_code', '')).strip()
        if code in EXCLUDED_DEMO_CODES: continue
        name = str(d.get('apt_name',  '')).strip()
        if not code or not name:
            continue

        hh       = safe_int(d.get('세대수', 0))
        dong     = safe_int(d.get('동수', 0))
        app_u    = safe_int(d.get('앱가입자수', 0))
        app_rate = safe_float(d.get('전체가입율', None))

        # 크기 구분 (단지(2번째숫자) 또는 앱컬럼 폴백)
        size = str(d.get('단지(2번째숫자)', '') or d.get('단지 유형', '')).strip() or None

        # 공시가격
        gongsi = safe_int(d.get('공시가격_평균(원)', 0)) or None

        # 연식
        age_raw = str(d.get('연식(3번째숫자)', '') or d.get('연식', '')).strip()
        age_type = age_raw or None

        # 공시가격 구분
        gongsi_grade = str(d.get('공시가격구분', '')).strip() or None

        # 평균 연령
        avg_age = safe_float(d.get('평균연령', None))

        # 서비스 / 솔루션 (비어있지 않은 것만 '/' 연결)
        svc_parts = [
            str(d.get(f'계약서비스{i}', '')).strip()
            for i in range(1, 5)
            if str(d.get(f'계약서비스{i}', '')).strip()
        ]
        svc = '/'.join(svc_parts) or None

        sol_parts = [
            str(d.get(f'솔루션구분{i}', '')).strip()
            for i in range(1, 5)
            if str(d.get(f'솔루션구분{i}', '')).strip()
        ]
        sol = '/'.join(sol_parts) or None

        region   = str(d.get('권역구분',  '')).strip() or None
        sido     = str(d.get('시도',      '')).strip() or None
        ch       = str(d.get('계약주체',  '')).strip() or None
        ctype    = str(d.get('단지구분',  '')).strip() or None

        # 계약 기간
        def fmt_date(raw):
            m = re.search(r'(\d{4})-(\d{2})-(\d{2})', str(raw))
            return f"{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else None

        contract_start = fmt_date(d.get('총계약시작일', ''))
        contract_end   = fmt_date(d.get('총계약종료일', ''))

        # 사용승인 연도
        build_yr_raw = str(d.get('사용승인일', '')).strip()
        build_yr_m   = re.search(r'(\d{4})', build_yr_raw)
        build_yr     = int(build_yr_m.group(1)) if build_yr_m else None

        # MAU 관련
        mau       = get_latest_mau(code)
        latest_ym = get_latest_ym(code)
        ratio     = round(mau / hh * 100, 1) if hh > 0 else None
        status    = classify_status(ratio)

        # 사업자정보로 보완
        biz = biz_map.get(code, {})

        row = {
            'code':           code,
            'name':           name,
            'hh':             hh,
            'dong':           dong,
            'app_users':      app_u,
            'app_rate':       app_rate,
            'size':           size,
            'gongsi':         gongsi,
            'age_type':       age_type,
            'age':            age_type,
            'gongsi_grade':   gongsi_grade,
            'avg_age':        avg_age,
            'svc':            svc,
            'sol':            sol,
            'region':         region,
            'sido':           sido,
            'contract_start': contract_start,
            'contract_end':   contract_end,
            'build_yr':       build_yr,
            'mau':            mau,
            'ratio':          ratio,
            'status':         status,
            'latest_ym':      latest_ym,
            'ch':             ch,
            'type':           ctype,
            'email':          biz.get('email'),
            'tel':            biz.get('tel'),
        }
        table.append(row)
        autocomplete.append({'n': name, 'c': code})

    # autocomplete 가나다 정렬
    autocomplete.sort(key=lambda x: x['n'])
    return table, autocomplete


# ═══════════════════════════════════════════════════════════════
# 6. svc_usage 빌드 (Rawdata 기반 + 실제 이용 이력 합산)
# ═══════════════════════════════════════════════════════════════
def build_svc_usage(table, vote_svc_cnt, vote_svc_refs,
                    survey_svc_cnt, survey_svc_refs, notify_svc_cnt):
    svc_matrix = {}
    for row in table:
        code = row['code']
        v_cnt  = vote_svc_cnt.get(code, 0)
        v_refs = len(vote_svc_refs.get(code, set()))
        s_cnt  = survey_svc_cnt.get(code, 0)
        s_refs = len(survey_svc_refs.get(code, set()))
        n_cnt  = notify_svc_cnt.get(code, 0)
        svc_matrix[code] = {
            'name':        row['name'],
            'vote':        v_cnt > 0,
            'survey':      s_cnt > 0,
            'notify':      n_cnt > 0,
            'svc_count':   (1 if v_cnt > 0 else 0) + (1 if s_cnt > 0 else 0) + (1 if n_cnt > 0 else 0),
            'vote_cnt':    v_cnt,
            'vote_refs':   v_refs,
            'survey_cnt':  s_cnt,
            'survey_refs': s_refs,
            'notify_cnt':  n_cnt,
        }
    cnt1 = sum(1 for v in svc_matrix.values() if v['svc_count'] >= 1)
    cnt2 = sum(1 for v in svc_matrix.values() if v['svc_count'] >= 2)
    cnt3 = sum(1 for v in svc_matrix.values() if v['svc_count'] >= 3)
    return {
        'svc_matrix': svc_matrix,
        'svc_counts': {'cnt1': cnt1, 'cnt2': cnt2, 'cnt3': cnt3},
    }


# ══════════════════════════════════════════════════════════════
# 6b. 누적 단지수 시계열 (1가지/2가지/3가지 이상 이용 누적)
# ═══════════════════════════════════════════════════════════════
def build_svc_count_history(ss):
    """단지별 첫 이용 일자 기준 월/주별 누적 cnt1/2/3 시계열."""
    first_use = {'vote': {}, 'notify': {}, 'survey': {}}
    for svc, sheet_name in [('vote', '전자투표'), ('notify', '알림서비스'), ('survey', '설문조사')]:
        header, rows = read_tab(ss, sheet_name)
        for d in to_dicts(header, rows):
            date_str = str(d.get('일자', '')).strip()[:10]
            code = str(d.get('단지코드', '')).strip()
            if code in EXCLUDED_DEMO_CODES: continue
            c = safe_int(d.get('건수', 0))
            if not date_str or not code or c <= 0:
                continue
            if not re.match(r'\d{4}-\d{2}-\d{2}', date_str):
                continue
            if code not in first_use[svc] or date_str < first_use[svc][code]:
                first_use[svc][code] = date_str

    all_codes = set()
    for svc in first_use:
        all_codes.update(first_use[svc].keys())

    all_dates = set()
    for svc in first_use:
        all_dates.update(first_use[svc].values())

    if not all_dates:
        return {
            'monthly_labels': [], 'cnt1_monthly': [], 'cnt2_monthly': [], 'cnt3_monthly': [],
            'weekly_labels': [], 'cnt1_weekly': [], 'cnt2_weekly': [], 'cnt3_weekly': [],
        }

    # 월별 누적
    monthly_points = sorted(set(dts[:7] for dts in all_dates))
    cnt1_m, cnt2_m, cnt3_m = [], [], []
    for t in monthly_points:
        c1 = c2 = c3 = 0
        for code in all_codes:
            n = 0
            for svc in first_use:
                if code in first_use[svc] and first_use[svc][code][:7] <= t:
                    n += 1
            if n >= 1: c1 += 1
            if n >= 2: c2 += 1
            if n >= 3: c3 += 1
        cnt1_m.append(c1); cnt2_m.append(c2); cnt3_m.append(c3)

    # 주별 누적 (ISO week)
    def to_iso_week(date_str):
        d = datetime.strptime(date_str, '%Y-%m-%d')
        yr, wn, _ = d.isocalendar()
        return f"{yr}-W{wn:02d}"

    weekly_points = sorted(set(to_iso_week(dts) for dts in all_dates))
    first_use_week = {svc: {c: to_iso_week(d) for c, d in first_use[svc].items()} for svc in first_use}

    cnt1_w, cnt2_w, cnt3_w = [], [], []
    for t in weekly_points:
        c1 = c2 = c3 = 0
        for code in all_codes:
            n = 0
            for svc in first_use_week:
                if code in first_use_week[svc] and first_use_week[svc][code] <= t:
                    n += 1
            if n >= 1: c1 += 1
            if n >= 2: c2 += 1
            if n >= 3: c3 += 1
        cnt1_w.append(c1); cnt2_w.append(c2); cnt3_w.append(c3)

    return {
        'monthly_labels': monthly_points,
        'cnt1_monthly': cnt1_m, 'cnt2_monthly': cnt2_m, 'cnt3_monthly': cnt3_m,
        'weekly_labels': weekly_points,
        'cnt1_weekly': cnt1_w, 'cnt2_weekly': cnt2_w, 'cnt3_weekly': cnt3_w,
    }




# ═══════════════════════════════════════════════════════════════
# 7. 성장률 계산
# ═══════════════════════════════════════════════════════════════
def pct(new, old):
    if old is None or old == 0:
        return None
    return round((new - old) / old * 100, 1)


def growth_series(values):
    n = len(values)
    mom = [None] * n
    qoq = [None] * n
    yoy = [None] * n
    for i in range(n):
        if i >= 1:  mom[i] = pct(values[i], values[i-1])
        if i >= 3:  qoq[i] = pct(values[i], values[i-3])
        if i >= 12: yoy[i] = pct(values[i], values[i-12])
    return mom, qoq, yoy


# ═══════════════════════════════════════════════════════════════
# 7b. D.gr 계산 — 일/주/월/분기/연간 시계열 (Sheets 기반 자동화)
# ═══════════════════════════════════════════════════════════════

def _date_to_isoweek(date_str):
    """'YYYY-MM-DD' → 'YYYY-Www'"""
    d = datetime.strptime(date_str, '%Y-%m-%d')
    yr, wn, _ = d.isocalendar()
    return f"{yr}-W{wn:02d}"


def _agg_weekly(day_dict):
    """{date_str: number} → sorted {YYYY-Www: number}"""
    wk = defaultdict(float)
    for ds, v in day_dict.items():
        if re.match(r'\d{4}-\d{2}-\d{2}', str(ds)):
            wk[_date_to_isoweek(ds)] += v
    return dict(sorted(wk.items()))


def _wow(vals):
    """주간 WoW 성장률 리스트"""
    n, w = len(vals), [None] * len(vals)
    for i in range(1, n):
        w[i] = pct(vals[i], vals[i - 1])
    return w


def _build_svc_gr(svc, day_cnt, day_rev, excl_rows, excl_sms, settled, excl_cmap=None):
    """
    한 서비스(vote/survey/notify)의 D.gr 키 전체 계산.
    day_cnt  : {date_str: row_count}
    day_rev  : {date_str: 추정매출(천원)}
    excl_rows: {ym: 무료단지 행 수}
    excl_sms : {ym: 무료단지 건수 합계}
    excl_cmap: {ym: {code: {name,rows,sms}}} — 단지별 제외 현황
    settled  : {ym: 정산금액(천원)}
    """
    if excl_cmap is None:
        excl_cmap = {}
    g = {}
    p = svc + '_'

    # ── 일간 ────────────────────────────────────────────────
    d_labels = sorted(day_cnt)
    d_vals   = [int(day_cnt[ds]) for ds in d_labels]
    d_rev    = [round(day_rev.get(ds, 0), 2) for ds in d_labels]
    g[p + 'day_labels']   = d_labels
    g[p + 'day_vals']     = d_vals
    g[p + 'rev_day_vals'] = d_rev

    # ── 주간 ────────────────────────────────────────────────
    wk_cnt = _agg_weekly(day_cnt)
    wk_rev = _agg_weekly(day_rev)

    wk_labels = sorted(wk_cnt)
    wk_vals   = [int(round(wk_cnt[w])) for w in wk_labels]
    g[p + 'wk_labels'] = wk_labels
    g[p + 'wk_vals']   = wk_vals
    g[p + 'wk_wow']    = _wow(wk_vals)

    rv_wk_l = [w for w in wk_labels if round(wk_rev.get(w, 0), 4) > 0]
    rv_wk_v = [round(wk_rev[w], 2) for w in rv_wk_l]
    g[p + 'rev_wk_labels'] = rv_wk_l
    g[p + 'rev_wk_vals']   = rv_wk_v
    g[p + 'rev_wk_wow']    = _wow(rv_wk_v)

    # ── 월간 ────────────────────────────────────────────────
    mo_cnt = defaultdict(int)
    mo_rev = defaultdict(float)
    for ds, v in day_cnt.items():
        mo_cnt[ds[:7]] += v
    for ds, v in day_rev.items():
        mo_rev[ds[:7]] += v

    mo_labels = sorted(mo_cnt)
    mo_vals   = [int(mo_cnt[m]) for m in mo_labels]
    mo_mom, mo_qoq, mo_yoy = growth_series(mo_vals)
    g[p + 'mo_labels'] = mo_labels
    g[p + 'mo_vals']   = mo_vals
    g[p + 'mo_mom']    = mo_mom
    g[p + 'mo_qoq']    = mo_qoq
    g[p + 'mo_yoy']    = mo_yoy

    rev_months = mo_labels
    rev_vals   = [round(mo_rev[m], 2) for m in rev_months]
    rv_mom, rv_qoq, rv_yoy = growth_series(rev_vals)
    g[p + 'rev_months'] = rev_months
    g[p + 'rev_vals']   = rev_vals
    g[p + 'rev_mom']    = rv_mom
    g[p + 'rev_qoq']    = rv_qoq
    g[p + 'rev_yoy']    = rv_yoy

    # ── 정산 분배 (월 정산금액 → 일/주 추정매출 비율로 안분) ──
    settle_ms = sorted(m for m in settled if m >= SETTLE_START and m in mo_rev and mo_rev[m] > 0)
    g[p + 'settle_labels'] = settle_ms
    g[p + 'settle_vals']   = [round(settled[m], 2) for m in settle_ms]

    s_day_l, s_day_v = [], []
    s_wk_d = defaultdict(float)
    for m in settle_ms:
        mo_tot = mo_rev[m]
        amt    = settled[m]
        for ds in d_labels:
            if ds[:7] != m:
                continue
            ratio = day_rev.get(ds, 0) / mo_tot
            sv    = round(amt * ratio, 2)
            s_day_l.append(ds)
            s_day_v.append(sv)
            s_wk_d[_date_to_isoweek(ds)] += sv

    g[p + 'settle_day_labels'] = s_day_l
    g[p + 'settle_day_vals']   = s_day_v

    s_wk_l = sorted(s_wk_d)
    s_wk_v = [round(s_wk_d[w], 2) for w in s_wk_l]
    g[p + 'settle_wk_labels'] = s_wk_l
    g[p + 'settle_wk_vals']   = s_wk_v
    g[p + 'settle_wk_wow']    = _wow(s_wk_v)

    # ── 무료단지 제외 현황 ───────────────────────────────────
    # excl_*: 정산 완료 월만
    excl_ms = [m for m in settle_ms if excl_sms.get(m, 0) > 0]
    g[p + 'excl_labels'] = excl_ms
    g[p + 'excl_cnt']    = [excl_sms.get(m, 0)            for m in excl_ms]   # SMS 발송 건수
    g[p + 'excl_cplx']   = [len(excl_cmap.get(m, {}))     for m in excl_ms]
    g[p + 'excl_amt']    = [sum(info.get('rev', 0) for info in excl_cmap.get(m, {}).values()) for m in excl_ms]  # 원 금액

    # excl_est_*: 전체 월
    all_ym = sorted(set(mo_labels) | set(excl_sms))
    g[p + 'excl_est_labels'] = all_ym
    g[p + 'excl_est_cnt']    = [excl_sms.get(m, 0)        for m in all_ym]
    g[p + 'excl_est_cplx']   = [len(excl_cmap.get(m, {})) for m in all_ym]
    g[p + 'excl_est_amt']    = [sum(info.get('rev', 0) for info in excl_cmap.get(m, {}).values()) for m in all_ym]

    return g


def compute_gr(v_day_cnt, v_day_rev, v_excl_rows, v_excl_sms, v_excl_cmap,
               s_day_cnt, s_day_rev, s_excl_rows, s_excl_sms, s_excl_cmap,
               n_day_cnt, n_day_rev, n_excl_rows, n_excl_sms, n_excl_cmap,
               settled):
    """D.gr 전체 계산 — 일/주/월/분기/연간 + 정산분배 + 미발행 제외"""
    gr = {}

    # 서비스별 시계열 — settled는 {svc: {ym: 만원}} 구조
    for svc, dc, dr, er, es, ecmap in [
        ('vote',   v_day_cnt, v_day_rev, v_excl_rows, v_excl_sms, v_excl_cmap),
        ('survey', s_day_cnt, s_day_rev, s_excl_rows, s_excl_sms, s_excl_cmap),
        ('notify', n_day_cnt, n_day_rev, n_excl_rows, n_excl_sms, n_excl_cmap),
    ]:
        gr.update(_build_svc_gr(svc, dc, dr, er, es, settled.get(svc, {}), ecmap))

    # excl_detail: 정산 완료 월 × 단지별 실적 (vote/survey/notify 모두 포함)
    all_settled_ym = set()
    for s_dict in settled.values():
        all_settled_ym.update(s_dict.keys())
    settle_ms = sorted(m for m in all_settled_ym if m >= SETTLE_START)
    excl_det  = []
    for m in settle_ms:
        for svc, cmap in [('vote', v_excl_cmap), ('survey', s_excl_cmap), ('notify', n_excl_cmap)]:
            for code, info in cmap.get(m, {}).items():
                excl_det.append({
                    'mo': m, 'code': code, 'name': info['name'],
                    'svc': svc, 'amt': info.get('rev', 0)
                })
    gr['excl_detail'] = excl_det

    # excl_est_detail: 전체 월 × 단지별 추정 (vote/survey/notify 모두 포함)
    all_excl_ym = sorted(set(list(v_excl_cmap) + list(s_excl_cmap) + list(n_excl_cmap)))
    excl_est = []
    for m in all_excl_ym:
        for svc, cmap in [('vote', v_excl_cmap), ('survey', s_excl_cmap), ('notify', n_excl_cmap)]:
            for code, info in cmap.get(m, {}).items():
                excl_est.append({
                    'mo': m, 'code': code, 'name': info['name'],
                    'svc': svc, 'amt': info.get('rev', 0)
                })
    gr['excl_est_detail'] = excl_est

    return gr


# ═══════════════════════════════════════════════════════════════
# 8. 전체 D 객체 조합
# ═══════════════════════════════════════════════════════════════

def build_cplx_top_data(ss, table):
    """단지별 서비스 사용량 top 200 + 매출 데이터 수집.
    Returns: {vote: [{code, name, hh, monthly:{ym:cnt}, rev_monthly:{ym:amt}, total_cnt, total_rev}], notify: [...], survey: [...]}
    """
    # 단지 정보 lookup (table에서)
    cplx_info = {}
    for t in (table or []):
        c = t.get('code', '') if isinstance(t, dict) else ''
        if c:
            cplx_info[c] = {'name': t.get('name', '') or c, 'hh': t.get('hh', 0) or 0}
    # 정산현황 — 단지별 월별 매출 (만원 단위)
    settle_by_cplx = {'vote': {}, 'notify': {}, 'survey': {}}
    SVC_MAP = {'전자투표': 'vote', '전자투표 (30% 할인)': 'vote', '알림서비스': 'notify', '설문조사': 'survey'}
    s_header, s_rows = read_tab(ss, '정산현황')
    if s_header:
        for d in to_dicts(s_header, s_rows):
            issued = str(d.get('발행여부', '')).strip().lower()
            if issued in ('cancel', 'void', '취소', '무효', 'skip'):
                continue
            svc = SVC_MAP.get(str(d.get('서비스명', '')).strip())
            if not svc: continue
            code = str(d.get('단지코드', '')).strip()
            if code in EXCLUDED_DEMO_CODES: continue
            if not code: continue
            ym_raw = str(d.get('정산월', '')).strip()
            m = re.match(r'(\d{4})[./-](\d{1,2})', ym_raw)
            if not m: continue
            ym = m.group(1) + '-' + str(int(m.group(2))).zfill(2)
            try:
                amt = float(str(d.get('청구금액', 0)).replace(',', '').replace('원', '').strip())
            except:
                continue
            settle_by_cplx[svc].setdefault(code, {}).setdefault(ym, 0.0)
            settle_by_cplx[svc][code][ym] += amt / 10000
    # 각 서비스 raw 시트 → 단지별 월별 cnt
    result = {}
    for svc, sheet_name in [('vote', '전자투표'), ('notify', '알림서비스'), ('survey', '설문조사')]:
        header, rows = read_tab(ss, sheet_name)
        cplx_monthly = {}  # code -> {ym: cnt}
        cplx_name = {}
        if header:
            for d in to_dicts(header, rows):
                date_str = str(d.get('일자', '')).strip()[:10]
                ym = ym_of(date_str)
                code = str(d.get('단지코드', '')).strip()
                if code in EXCLUDED_DEMO_CODES: continue
                name = str(d.get('단지명', '')).strip()
                c = safe_int(d.get('건수', 0))
                if not ym or not code or c <= 0: continue
                cplx_monthly.setdefault(code, {}).setdefault(ym, 0)
                cplx_monthly[code][ym] += c
                if name and code not in cplx_name:
                    cplx_name[code] = name
        # top 200 (누적 cnt 기준)
        totals = [(code, sum(mo.values())) for code, mo in cplx_monthly.items()]
        totals.sort(key=lambda x: -x[1])
        top200 = totals[:200]
        out = []
        for code, total_cnt in top200:
            info = cplx_info.get(code, {})
            name = info.get('name') or cplx_name.get(code, code)
            hh = info.get('hh', 0)
            rev_m = {ym: round(amt, 2) for ym, amt in settle_by_cplx[svc].get(code, {}).items()}
            total_rev = round(sum(rev_m.values()), 2)
            out.append({
                'code': code,
                'name': name,
                'hh': hh,
                'monthly': {ym: int(cnt) for ym, cnt in cplx_monthly[code].items()},
                'rev_monthly': rev_m,
                'total_cnt': int(total_cnt),
                'total_rev': total_rev
            })
        result[svc] = out
        print('  cplx_top.' + svc + ': ' + str(len(out)) + '개 단지')
    return result



def build_yearly_compare(D_obj):
    """년도별 월별 매출 비교."""
    months_all = D_obj.get('months', [])
    years_set = set()
    for m in months_all:
        if isinstance(m, str) and len(m) >= 4:
            years_set.add(m[:4])
    years = sorted(years_set)
    gr = D_obj.get('gr', {})
    def gm(labels, vals):
        out = {}
        for i, lab in enumerate(labels or []):
            if not isinstance(lab, str) or len(lab) < 7: continue
            out[(lab[:4], int(lab[5:7]))] = float((vals or [0])[i] if i < len(vals) else 0)
        return out
    apt_est, apt_st_excl, apt_pp, of_data, as_data = {}, {}, {}, {}, {}
    for svc in ['vote', 'notify', 'survey']:
        for k, v in gm(gr.get(svc + '_rev_months', []) or gr.get(svc + '_mo_labels', []), gr.get(svc + '_rev_vals', [])).items():
            apt_est[k] = apt_est.get(k, 0) + v
        for k, v in gm(gr.get(svc + '_settle_labels', []), gr.get(svc + '_settle_vals', [])).items():
            apt_st_excl[k] = apt_st_excl.get(k, 0) + v
    pp = D_obj.get('prepaid_promo', {})
    for k, v in gm(pp.get('months', []), pp.get('values', [])).items():
        apt_pp[k] = apt_pp.get(k, 0) + v
    of = D_obj.get('officener', {})
    for k, v in gm(of.get('months', []), of.get('values', [])).items():
        of_data[k] = v
    as_ = D_obj.get('aptstory', {})
    for k, v in gm(as_.get('months', []), as_.get('values', [])).items():
        as_data[k] = v
    def to_arr(d):
        return {y: [round(d.get((y, m), 0), 2) for m in range(1, 13)] for y in years}
    apt_st = {k: v + apt_pp.get(k, 0) for k, v in apt_st_excl.items()}
    total_est = {}
    for d in (apt_est, of_data, as_data):
        for k, v in d.items(): total_est[k] = total_est.get(k, 0) + v
    total_st = {}
    for d in (apt_st, of_data, as_data):
        for k, v in d.items(): total_st[k] = total_st.get(k, 0) + v
    return {
        'years': years,
        'data': {
            'aptner_est': to_arr(apt_est), 'aptner_st': to_arr(apt_st),
            'officener_est': to_arr(of_data), 'officener_st': to_arr(of_data),
            'aptstory_est': to_arr(as_data), 'aptstory_st': to_arr(as_data),
            'total_est': to_arr(total_est), 'total_st': to_arr(total_st)
        }
    }


def build_free_cplx_data(ss, free, table):
    """전자투표 무료단지."""
    if not free:
        return []
    cplx_info = {}
    for t in (table or []):
        c = t.get('code', '') if isinstance(t, dict) else ''
        if c:
            cplx_info[c] = {'name': t.get('name', '') or c, 'hh': t.get('hh', 0) or 0}
    header, rows = read_tab(ss, '전자투표')
    cplx_monthly, cplx_weekly, cplx_name = {}, {}, {}
    if header:
        for d in to_dicts(header, rows):
            date_str = str(d.get('일자', '')).strip()[:10]
            ym = ym_of(date_str)
            wk = _date_to_isoweek(date_str)
            code = str(d.get('단지코드', '')).strip()
            if code in EXCLUDED_DEMO_CODES: continue
            name = str(d.get('단지명', '')).strip()
            c = safe_int(d.get('건수', 0))
            if not ym or not code or c <= 0: continue
            if code not in free: continue
            cplx_monthly.setdefault(code, {}).setdefault(ym, 0)
            cplx_monthly[code][ym] += c
            cplx_weekly.setdefault(code, {}).setdefault(wk, 0)
            cplx_weekly[code][wk] += c
            if name and code not in cplx_name:
                cplx_name[code] = name
    out = []
    for code in cplx_monthly:
        monthly = cplx_monthly[code]
        weekly = cplx_weekly.get(code, {})
        monthly_est = {ym: round(cnt * 44 / 10000, 2) for ym, cnt in monthly.items()}
        info = cplx_info.get(code, {})
        out.append({
            'code': code,
            'name': info.get('name') or cplx_name.get(code, code),
            'hh': info.get('hh', 0),
            'monthly_cnt': {ym: int(cnt) for ym, cnt in monthly.items()},
            'weekly_cnt': {wk: int(cnt) for wk, cnt in weekly.items()},
            'monthly_est': monthly_est,
            'total_cnt': int(sum(monthly.values()))
        })
    out.sort(key=lambda x: -x['total_cnt'])
    print('  free_cplx: ' + str(len(out)) + '개 무료단지')
    return out


def build_d(ss):
    print("[1/6] 참조 탭 읽기...")
    free    = load_code_set(ss, '무료단지')
    jikbang = load_jikbang(ss)           # set of (code, ym) — 기준월 기준
    promo_snap = load_promo(ss)          # {(code, ym): {remain, startDate}} - sheet snapshot
    promo   = load_promo_legacy(ss)      # {code: 1.0} for aggregate_* (raw revenue est.)
    discount_30_codes = load_discount_30(ss)
    officener_rev = load_officener_rev(ss)
    aptstory_rev = load_aptstory_rev(ss)
    prepaid_promo = load_prepaid_promo(ss)
    print(f"  오피스너 {len(officener_rev['months'])}개월 / 아파트스토리 {len(aptstory_rev['months'])}개월 / 선결제 {len(prepaid_promo['months'])}개월")
    # settlement: 시트 '정산현황' 청구금액 컬럼 합계 (5월 외부 KPI와 일치 검증)
    settled = load_settlement(ss)
    settled_mo_counts = {svc: len(d) for svc, d in settled.items()}
    jikbang_cplx = len(set(c for c, _ in jikbang))
    print(f"  무료 {len(free)}개 / 직방전환 {jikbang_cplx}개단지·{len(jikbang)}건(기준월) / 프로모션 {len(promo)}개 / 정산 vote={settled_mo_counts['vote']}개월 survey={settled_mo_counts['survey']}개월 notify={settled_mo_counts['notify']}개월")

    print("[2/6] 원본 탭 집계...")
    v_cnt, v_ids, v_rev, v_cplx_ids, v_svc_cnt, v_svc_refs, \
        v_day_cnt, v_day_rev, v_excl_rows, v_excl_sms, v_excl_cmap = aggregate_vote(ss, free, jikbang, promo)
    s_cnt, s_ids, s_rev, s_cplx_ids, s_svc_cnt, s_svc_refs, \
        s_day_cnt, s_day_rev, s_excl_rows, s_excl_sms, s_excl_cmap = aggregate_survey(ss, free, promo)
    n_snd, n_ids, n_rev, n_ev, n_svc_cnt, n_day_cnt, n_day_rev, \
        n_excl_rows, n_excl_sms, n_excl_cmap                       = aggregate_notify(ss, free)
    print(f"  vote {sum(v_cnt.values()):,}건 / survey {sum(s_cnt.values()):,}건 / notify {sum(n_snd.values()):,}건")
    print(f"  일간 vote {len(v_day_cnt)}일 / survey {len(s_day_cnt)}일 / notify {len(n_day_cnt)}일")

    # ── 공통 월 레이블 ────────────────────────────────────────
    all_months_set = set(v_cnt) | set(s_cnt) | set(n_snd)
    all_months     = sorted(all_months_set)
    vote_months    = sorted(v_cnt)   # 전자투표 있는 달

    # ── 전자투표 시계열 ───────────────────────────────────────
    vote_ids_values = [len(v_ids.get(m, set())) for m in vote_months]
    vote_cnt_values = [v_cnt.get(m, 0)          for m in vote_months]

    # ── 알림서비스 시계열 ─────────────────────────────────────
    notify_labels = sorted(n_snd)
    notify_events = [len(n_ev.get(m, set()))  for m in notify_labels]
    notify_sends  = [n_snd.get(m, 0)          for m in notify_labels]

    # ── 설문조사 시계열 ───────────────────────────────────────
    survey_labels     = sorted(s_cnt)
    survey_ids_values = [len(s_ids.get(m, set())) for m in survey_labels]
    survey_cnt_values = [s_cnt.get(m, 0)          for m in survey_labels]

    # ── 매출 추정 ─────────────────────────────────────────────
    rev_labels = all_months
    rev_values = [round(v_rev.get(m, 0) + s_rev.get(m, 0) + n_rev.get(m, 0), 2) for m in rev_labels]

    # 전자투표 전용 매출 추정
    vote_rev_est_values = [round(v_rev.get(m, 0), 2) for m in rev_labels]

    # ── 정산 매출 ─────────────────────────────────────────────
    # 서비스별 settled 합산하여 전체 월별 정산 합계 생성
    settled_total = {}
    for svc_dict in settled.values():
        for ym, amt in svc_dict.items():
            settled_total[ym] = round(settled_total.get(ym, 0) + amt, 2)

    rev_settled_values = []
    rev_settled_types  = []
    for i, m in enumerate(rev_labels):
        if m in settled_total and m >= SETTLE_START:
            rev_settled_values.append(settled_total[m])
            rev_settled_types.append('settled')
        else:
            rev_settled_values.append(rev_values[i])
            rev_settled_types.append('estimated')

    # 전자투표 전용 정산 매출
    vote_settled_dict = settled.get('vote', {})
    vote_rev_settled_values = []
    for i, m in enumerate(rev_labels):
        if m in vote_settled_dict and m >= SETTLE_START:
            vote_rev_settled_values.append(vote_settled_dict[m])
        else:
            vote_rev_settled_values.append(vote_rev_est_values[i])

    # ── 성장률 ────────────────────────────────────────────────
    mom_vote, qoq_vote, yoy_vote = growth_series(vote_cnt_values)
    vote_ids_mom  = [None] + [pct(vote_ids_values[i], vote_ids_values[i-1]) for i in range(1, len(vote_ids_values))]
    notify_ev_mom  = [None] + [pct(notify_events[i], notify_events[i-1])    for i in range(1, len(notify_events))]
    notify_snd_mom = [None] + [pct(notify_sends[i],  notify_sends[i-1])     for i in range(1, len(notify_sends))]
    survey_ids_mom = [None] + [pct(survey_ids_values[i], survey_ids_values[i-1]) for i in range(1, len(survey_ids_values))]

    # ── 단지수 (svc_cplx) ─────────────────────────────────────
    svc_cplx = {
        'vote_cplx_monthly':   {m: len(v_cplx_ids.get(m, set())) for m in vote_months},
        'notify_cplx_monthly': {m: len(n_ids.get(m, set()))      for m in notify_labels},
        'survey_cplx_monthly': {m: len(s_cplx_ids.get(m, set())) for m in survey_labels},
    }

    print("[3/6] 단지별 MAU 로드...")
    mau_m, top20_mau = load_mau(ss, all_months)
    print(f"  {len(mau_m)}개 단지, {len(all_months)}개월 기준 top20 생성")

    print("[4/6] Rawdata → table / autocomplete / svc_usage ...")
    table, autocomplete = load_table_and_autocomplete(ss, mau_m)
    svc_usage = build_svc_usage(
        table, v_svc_cnt, v_svc_refs, s_svc_cnt, s_svc_refs, n_svc_cnt
    )
    svc_count_history = build_svc_count_history(ss)
    print(f"  누적 단지수: 월 {len(svc_count_history['monthly_labels'])}개 / 주 {len(svc_count_history['weekly_labels'])}개")
    print(f"  단지 {len(table)}개 / autocomplete {len(autocomplete)}개 / svc_matrix {len(svc_usage['svc_matrix'])}개")

    print("[4b/6] D.gr 계산 (일/주/월/분기/연간)...")
    gr = compute_gr(
        v_day_cnt, v_day_rev, v_excl_rows, v_excl_sms, v_excl_cmap,
        s_day_cnt, s_day_rev, s_excl_rows, s_excl_sms, s_excl_cmap,
        n_day_cnt, n_day_rev, n_excl_rows, n_excl_sms, n_excl_cmap,
        settled
    )
    gr_keys = len([k for k in gr if not k.startswith('excl')])
    print(f"  gr 필드 {len(gr)}개 (서비스별 시계열 {gr_keys}개 + 제외현황 포함)")

    D = {
        # 월 레이블
        'months':     vote_months,
        'all_months': all_months,
        # MAU top20
        'top20_mau':  top20_mau,
        # 전자투표
        'vote_ids_labels':  vote_months,
        'vote_ids_values':  vote_ids_values,
        'vote_cnt_values':  vote_cnt_values,
        # 알림서비스
        'notify_labels': notify_labels,
        'notify_events': notify_events,
        'notify_sends':  notify_sends,
        # 설문조사
        'survey_labels':     survey_labels,
        'survey_ids_values': survey_ids_values,
        'survey_cnt_values': survey_cnt_values,
        # 매출 (전체 합산 - 차트용)
        'revenue_estimate_labels': rev_labels,
        'revenue_estimate_values': rev_values,
        'revenue_settled_values':  rev_settled_values,
        'revenue_settled_types':   rev_settled_types,
        # 전자투표 전용 매출
        'vote_rev_est_values':     vote_rev_est_values,
        'vote_rev_settled_values': vote_rev_settled_values,
        # 성장률
        'growth_labels':   vote_months,
        'mom_vote':        mom_vote,
        'qoq_vote':        qoq_vote,
        'yoy_vote':        yoy_vote,
        'notify_ev_mom':   notify_ev_mom,
        'notify_snd_mom':  notify_snd_mom,
        'vote_ids_mom':    vote_ids_mom,
        'survey_ids_mom':  survey_ids_mom,
        # 단지수
        'svc_cplx': svc_cplx,
        # 단지 데이터
        'table':       table,
        'autocomplete': autocomplete,
        'mau_m':       mau_m,
        'svc_usage':   svc_usage,
        'svc_count_history': svc_count_history,
        'cplx_top': build_cplx_top_data(ss, table),
        'free_cplx': build_free_cplx_data(ss, free, table),
        # 날짜
        'data_as_of': datetime.now().strftime('%Y-%m-%d'),
        # 일/주/월/분기/연간 시계열 (D.gr — Sheets 기반 자동 갱신)
        'gr': gr,
        # 오피스너 / 아파트스토리 매출 시계열
        'officener': officener_rev,
        'aptstory': aptstory_rev,
        'prepaid_promo': prepaid_promo,
    }
    D['yearly_compare'] = build_yearly_compare(D)
    return D


# ═══════════════════════════════════════════════════════════════
# 9. 기존 HTML 에서 안정 필드 추출 & 주입
# ═══════════════════════════════════════════════════════════════
STABLE_KEYS = [
    # 앱 활동 데이터 (Sheets에 없음 — 기존 HTML에서 보존)
    'complex_activity_top200',
    'activity_dist_labels', 'activity_dist_values', 'activity_dist_colors',
    'activity_monthly_months', 'activity_monthly_series',
    # ※ 'gr' 제거 — compute_gr()로 Sheets 기반 자동 계산
]


def find_latest_html():
    # SOURCE_VERSION fixed as base (e.g. v113); next save = max_ver+1
    files = glob.glob(os.path.join(BASE_DIR, f'{HTML_PREFIX}*.html'))
    versions = [(int(m.group(1)), f)
                for f in files
                for m in [re.search(r'_v(\d+)\.html$', f)] if m]
    if not versions:
        return None, 0
    src_path = next((f for v, f in versions if v == SOURCE_VERSION), None)
    max_ver = max(v for v, _ in versions)
    if src_path is None:
        versions.sort()
        return versions[-1][1], max_ver
    return src_path, max_ver


def extract_d_from_html(html_path):
    """var D = {...}; 한 줄을 JSON 파싱"""
    try:
        with open(html_path, encoding='utf-8') as f:
            for line in f:
                if re.match(r'^\s*var D\s*=\s*\{', line):
                    m = re.match(r'^\s*var D\s*=\s*(\{.*\});', line)
                    if m:
                        return json.loads(m.group(1))
    except Exception as e:
        print(f"  ⚠ D 객체 추출 실패: {e}")
    return {}


def _validate_template(path):
    """v113 템플릿 무결성 검증"""
    try:
        with open(path, 'rb') as f:
            raw = f.read()
    except Exception as e:
        return False, "읽기 실패: " + str(e)
    null_n = raw.count(b'\x00')
    if null_n > 0:
        return False, "null bytes " + str(null_n) + "개"
    if not raw.rstrip().endswith(b'</html>'):
        return False, "파일 끝 </html> 누락 (잘림 의심)"
    text = raw.decode('utf-8', errors='replace')
    n_open = text.count('<script')
    n_close = text.count('</script>')
    if n_open != n_close:
        return False, "<script> 불일치 (" + str(n_open) + " vs " + str(n_close) + ")"
    if len(raw) < 1000000:
        return False, "파일 크기 너무 작음 (" + str(len(raw)) + " bytes)"
    return True, "OK (size=" + str(len(raw)) + ", scripts=" + str(n_open) + ")"


def generate_html(D_new):
    print("[5/6] 기존 HTML → 안정 필드 추출...")
    src_path, cur_ver = find_latest_html()
    if not src_path:
        print("  ✗ 기존 HTML 없음")
        sys.exit(1)
    print(f"  소스: {os.path.basename(src_path)} (v{cur_ver})")
    # 🛡 템플릿 무결성 검증
    ok, msg = _validate_template(src_path)
    if not ok:
        print("  🚨 템플릿 무결성 실패: " + msg)
        print("  복구: git checkout HEAD -- \"" + os.path.basename(src_path) + "\"")
        sys.exit(1)
    print("  ✓ 템플릿 검증: " + msg)
    # 자동 백업
    bak_path = src_path + '.bak'
    try:
        with open(src_path, 'rb') as _s, open(bak_path, 'wb') as _d:
            _d.write(_s.read())
        print("  ✓ 백업: " + os.path.basename(bak_path))
    except Exception as e:
        print("  ⚠ 백업 실패: " + str(e))


    D_old = extract_d_from_html(src_path)
    for key in STABLE_KEYS:
        if key in D_old:
            D_new[key] = D_old[key]   # 항상 기존값으로 덮어쓰기
        else:
            print(f"  ⚠ 안정 필드 '{key}' 기존 HTML에 없음")

    print("[6/6] HTML 파일 생성...")
    next_ver  = cur_ver + 1
    out_name  = f"{HTML_PREFIX}{next_ver}.html"
    out_path  = os.path.join(BASE_DIR, out_name)
    idx_path  = os.path.join(BASE_DIR, 'index.html')

    d_json = json.dumps(D_new, ensure_ascii=False, separators=(',', ':'))

    with open(src_path, encoding='utf-8') as f:
        lines = f.readlines()

    replaced  = False
    new_lines = []
    for line in lines:
        if not replaced and re.match(r'^\s*var D\s*=\s*\{', line):
            new_lines.append(f'var D = {d_json};\n')
            replaced = True
        else:
            new_lines.append(line)

    if not replaced:
        print("  ✗ var D = ... 라인을 찾지 못했습니다.")
        sys.exit(1)

    html_content = ''.join(new_lines)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
        f.flush()
    # index.html — Python 레벨 바이트 복사 (한글 경로 안전)
    import os as _os
    with open(out_path, 'rb') as _src:
        _data = _src.read()
    with open(idx_path, 'wb') as _dst:
        _dst.write(_data)
    # 크기 검증
    sz_out = _os.path.getsize(out_path)
    sz_idx = _os.path.getsize(idx_path)
    if sz_out != sz_idx:
        print(f"  ⚠ index.html 크기 불일치 ({sz_idx} vs {sz_out}) — 재시도")
        with open(out_path, 'rb') as _src:
            _data = _src.read()
        with open(idx_path, 'wb') as _dst:
            _dst.write(_data)


    sz_kb = len(d_json) / 1024
    print(f"  저장: {out_name}  (D={sz_kb:.0f} KB)")
    print(f"  저장: index.html")
    return next_ver


if __name__ == '__main__':
    print("=== 단지서비스팀 대시보드 HTML 생성 시작 ===")
    gc = connect_sheets()
    ss = gc.open_by_key(SPREADSHEET_ID)
    D_new  = build_d(ss)
    new_ver = generate_html(D_new)
    print(f"\n[완료] v{new_ver} 생성")
    print("  다음: python deploy.py 로 배포")

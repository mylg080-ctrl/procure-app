import streamlit as st
import calendar
import json
import os
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote
import io
import time
import xml.etree.ElementTree as ET

import requests
from openpyxl import Workbook

# ───────────── 페이지 설정 ─────────────
st.set_page_config(
    page_title="공공조달 통합 낙찰결과 조회",
    page_icon="📊",
    layout="wide"
)

# ───────────── 공통 설정 ─────────────
G2B_BASE_URL = "https://apis.data.go.kr/1230000/as/ScsbidInfoService"
G2B_OPERATION = "getScsbidListSttusThngPPSSrch"
G2B_ROWS = 100
G2B_MAX_PER_MONTH = 3000
G2B_BID_BASE_URL = "https://apis.data.go.kr/1230000/ad/BidPublicInfoService"
G2B_OP_BSIS = "getBidPblancListInfoThngBsisAmount"
G2B_OP_NOTICE = "getBidPblancListInfoThng"

D2B_BASE_URL = "https://apis.data.go.kr/1690000/BidResultInfoService"
D2B_LIST_OP = "getDmstcCmpetBidResultList"
D2B_ROWS = 1000

SYSTEMS = {
    "g2b": dict(
        label="나라장터 (조달청)", file_prefix="나라장터",
        name_label="수요기관명", months=12, years=5,
    ),
    "d2b": dict(
        label="국방전자조달 (방위사업청)", file_prefix="D2B",
        name_label="부대명", months=12, years=5,
    ),
}

class Fatal(Exception):
    pass

QUOTA_MSG = "오늘 API 호출 한도를 초과했습니다. 공공데이터포털 개발계정 일일 호출 한도를 확인하세요."

def clean_key(key):
    key = (key or "").strip()
    if "%" in key:
        key = unquote(key)
    return key

def num(v):
    if v is None or str(v).strip() == "":
        return None
    try:
        f = float(str(v).replace(",", ""))
    except ValueError:
        return None
    return int(f) if f.is_integer() else f

def fmt_amt(v):
    return "" if v is None else f"{int(v):,}"

def fmt_rate(v):
    return "" if v is None else f"{v:.3f}"

def fmt_dt(v):
    d = re.sub(r"\D", "", v or "")
    if len(d) >= 12:
        return f"{d[:4]}-{d[4:6]}-{d[6:8]} {d[8:10]}:{d[10:12]}"
    if len(d) >= 8:
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"
    return v or ""

def fmt_biz(v):
    d = re.sub(r"\D", "", v or "")
    return f"{d[:3]}-{d[3:5]}-{d[5:]}" if len(d) == 10 else (v or "")

def month_windows(months):
    now = datetime.now()
    cur = now.year * 12 + now.month - 1
    out = []
    for idx in range(cur - months, cur + 1):
        y, m = divmod(idx, 12)
        m += 1
        last = calendar.monthrange(y, m)[1]
        end = now.strftime("%Y%m%d%H%M") if idx == cur else f"{y}{m:02d}{last}2359"
        out.append((f"{y}{m:02d}010000", end))
    return out[::-1]

def calc_gap(r):
    a, b = r.get("낙찰률(%)"), r.get("낙찰하한율(%)")
    r["GAP(%p)"] = round(a - b, 3) if a is not None and b is not None else None

# ───────────── G2B API ─────────────
def g2b_call(op, params, base=G2B_BASE_URL, retries=3):
    for attempt in range(retries):
        try:
            r = requests.get(f"{base}/{op}", params=params, timeout=60)
            r.raise_for_status()
            try:
                return r.json()
            except ValueError:
                head = r.text[:300]
                if "LIMITED_NUMBER" in head.upper():
                    raise Fatal(QUOTA_MSG)
                raise Fatal(f"API 응답 형식 오류\n{head}")
        except Fatal:
            raise
        except Exception:
            if attempt == retries - 1:
                raise Fatal("네트워크 오류 발생")
            time.sleep(1)

def g2b_parse(data):
    resp = data.get("response") or next((v for v in data.values() if isinstance(v, dict) and "header" in v), data)
    header = resp.get("header", {}) or {}
    body = resp.get("body") or {}
    items = body.get("items") or []
    if isinstance(items, dict):
        items = items.get("item", [])
    if isinstance(items, dict):
        items = [items]
    total = int(body.get("totalCount", 0) or 0)
    return str(header.get("resultCode", "")), header.get("resultMsg", ""), items, total

def g2b_fetch_month(key, bgn, end, div, code=None, name=None):
    rows, page = [], 1
    while True:
        params = {"ServiceKey": key, "type": "json", "pageNo": page, "numOfRows": G2B_ROWS,
                  "inqryDiv": div, "inqryBgnDt": bgn, "inqryEndDt": end}
        if code:
            params["dminsttCd"] = code
        if name:
            params["dminsttNm"] = name
        rc, msg, items, total = g2b_parse(g2b_call(G2B_OPERATION, params))
        if rc == "03":
            return []
        if rc != "00":
            raise Fatal(f"[{rc}] {msg}")
        if total > G2B_MAX_PER_MONTH:
            raise Fatal(f"한 달 결과가 너무 많습니다 ({total}건).")
        rows += items
        if not items or len(rows) >= total:
            break
        page += 1
    if code:
        rows = [r for r in rows if not r.get("dminsttCd") or r.get("dminsttCd") == code]
    return rows

def g2b_pick_ord(items, ord_):
    for it in items:
        if str(it.get("bidNtceOrd", "")) == str(ord_):
            return it
    return items[-1] if items else None

def g2b_fetch_notice_extra(key, no, ord_):
    base = {"ServiceKey": key, "type": "json", "pageNo": 1, "numOfRows": 10,
            "inqryDiv": "2", "bidNtceNo": no}
    out = {"bssamt": None, "lwlt": None, "mthd": ""}
    rc, _, items, _ = g2b_parse(g2b_call(G2B_OP_BSIS, dict(base), base=G2B_BID_BASE_URL))
    it = g2b_pick_ord(items, ord_)
    if it:
        out["bssamt"] = num(it.get("bssamt"))
    
    rc, _, items, _ = g2b_parse(g2b_call(G2B_OP_NOTICE, dict(base), base=G2B_BID_BASE_URL))
    it = g2b_pick_ord(items, ord_)
    if it:
        out["lwlt"] = num(it.get("sucsfbidLwltRate"))
        out["mthd"] = it.get("sucsfbidMthdNm") or ""
    return out

def g2b_normalize(it):
    no, od = it.get("bidNtceNo", ""), it.get("bidNtceOrd", "")
    return {
        "개찰일시": it.get("rlOpengDt") or "",
        "공고번호": no,
        "차수": od,
        "공고명": it.get("bidNtceNm") or "",
        "수요기관": it.get("dminsttNm") or "",
        "결과": "낙찰",
        "참가업체수": num(it.get("prtcptCnum")),
        "낙찰업체": it.get("bidwinnrNm") or "",
        "사업자번호": it.get("bidwinnrBizno") or "",
        "대표자": it.get("bidwinnrCeoNm") or "",
        "낙찰금액": num(it.get("sucsfbidAmt")),
        "낙찰률(%)": num(it.get("sucsfbidRate")),
        "기초금액": None,
        "낙찰하한율(%)": num(it.get("sucsfbidLwltRate")),
        "GAP(%p)": None,
        "낙찰방법": "",
        "예산액": None,
    }

# ───────────── D2B API ─────────────
def d2b_call(params, op, retries=3):
    for attempt in range(retries):
        try:
            r = requests.get(f"{D2B_BASE_URL}/{op}", params=params, timeout=30)
            root = ET.fromstring(r.text)
            err_msg = root.findtext(".//errMsg") or root.findtext(".//returnAuthMsg")
            if err_msg:
                if "LIMITED_NUMBER" in err_msg:
                    raise Fatal(QUOTA_MSG)
                raise Fatal(f"API 서버 오류:\n{err_msg}")
            return root
        except Fatal:
            raise
        except Exception:
            if attempt == retries - 1:
                raise Fatal("네트워크 또는 XML 파싱 오류")
            time.sleep(1)

def d2b_parse(root):
    rc = root.findtext(".//resultCode") or ""
    msg = root.findtext(".//resultMsg") or ""
    total = int(root.findtext(".//totalCount") or "0")
    items = [{child.tag: child.text for child in item} for item in root.findall(".//item")]
    return rc, msg, items, total

def d2b_item_inst(it):
    return it.get("orntCode") or "", it.get("ornt") or ""

def d2b_fetch_month(key, bgn, end, code=None):
    rows, page = [], 1
    while True:
        params = {"serviceKey": key, "pageNo": page, "numOfRows": D2B_ROWS,
                  "opengDateBegin": bgn[:8], "opengDateEnd": end[:8]}
        if code:
            params["orntCode"] = code
        rc, msg, items, total = d2b_parse(d2b_call(params, D2B_LIST_OP))
        if rc not in ("", "00", "0", "03"):
            raise Fatal(f"[{rc}] {msg}")
        if rc == "03" or not items:
            break
        rows += items
        if len(rows) >= total:
            break
        page += 1
    if code:
        rows = [r for r in rows if not d2b_item_inst(r)[0] or d2b_item_inst(r)[0] == code]
    return rows

def d2b_normalize(it):
    meth = " · ".join(x for x in (it.get("cntrctMth"), it.get("sucbidrDecsnMth")) if x)
    return {
        "개찰일시": fmt_dt(it.get("opengDt") or it.get("opengDate")),
        "공고번호": it.get("pblancNo") or "",
        "차수": it.get("pblancOdr") or "",
        "공고명": it.get("bidNm") or "",
        "수요기관": d2b_item_inst(it)[1],
        "결과": it.get("bidResult") or "",
        "참가업체수": None,
        "낙찰업체": it.get("scsbidEntrpsNm") or "",
        "사업자번호": fmt_biz(it.get("scsbidEntrpsBsnmRegistNo")),
        "대표자": "",
        "낙찰금액": num(it.get("scsbidAmnt")),
        "낙찰률(%)": num(it.get("scsbidRt")),
        "기초금액": num(it.get("bsisPreparPc")),
        "낙찰하한율(%)": num(it.get("scsbidLwltRt")),
        "GAP(%p)": None,
        "낙찰방법": meth,
        "예산액": num(it.get("budgetAmount")),
    }

def generate_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "개찰결과"
    cols = ["개찰일시", "공고번호", "차수", "공고명", "수요기관", "결과", "낙찰업체", "기초금액", "낙찰하한율(%)", "낙찰금액", "낙찰률(%)", "GAP(%p)", "낙찰방법"]
    ws.append(cols)
    for r in rows:
        ws.append([
            r.get("개찰일시"), r.get("공고번호"), r.get("차수"), r.get("공고명"), r.get("수요기관"), r.get("결과"), r.get("낙찰업체"),
            r.get("기초금액"), r.get("낙찰하한율(%)"), r.get("낙찰금액"), r.get("낙찰률(%)"), r.get("GAP(%p)"), r.get("낙찰방법")
        ])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ───────────── UI 구성 ─────────────
st.title("📊 공공조달 통합 낙찰결과 조회")

if "institutions" not in st.session_state:
    st.session_state.institutions = {}
if "rows" not in st.session_state:
    st.session_state.rows = []

with st.sidebar:
    st.header("⚙️ 설정 및 인증")
    sys_type = st.radio("조달 시스템 선택", ["g2b", "d2b"], format_func=lambda x: SYSTEMS[x]["label"])
    service_key = st.text_input("공공데이터포털 서비스키 (Decoding)", type="password")
    
    st.divider()
    st.subheader("① 기관 검색")
    inst_name = st.text_input(SYSTEMS[sys_type]["name_label"], placeholder="두 글자 이상 입력")
    search_months = st.number_input("검색 범위 (개월)", min_value=1, max_value=60, value=12)
    
    if st.button("기관 검색 실행"):
        key = clean_key(service_key)
        if not key:
            st.error("서비스키를 입력해주세요.")
        elif len(inst_name) < 2:
            st.warning("기관명을 두 글자 이상 입력해주세요.")
        else:
            with st.spinner("기관 검색 중..."):
                try:
                    found = {}
                    wins = month_windows(search_months)
                    for bgn, end in wins:
                        if sys_type == "g2b":
                            res = g2b_fetch_month(key, bgn, end, "2", name=inst_name)
                            for r in res:
                                cd, nm = r.get("dminsttCd"), r.get("dminsttNm") or ""
                                if cd and inst_name in nm:
                                    found[cd] = nm
                        else:
                            res = d2b_fetch_month(key, bgn, end)
                            for r in res:
                                cd, nm = d2b_item_inst(r)
                                if nm and inst_name in nm:
                                    found[cd or nm] = nm
                    st.session_state.institutions = found
                    st.success(f"검색 완료! 총 {len(found)}개 기관 발견")
                except Exception as e:
                    st.error(str(e))

col1, col2 = st.columns(2)

with col1:
    st.subheader("② 기관 선택 및 조회")
    inst_options = st.session_state.institutions
    if inst_options:
        selected_cd = st.selectbox("검색된 기관 선택", options=list(inst_options.keys()), format_func=lambda x: f"{inst_options[x]} ({x})")
    else:
        selected_cd = st.text_input("기관 코드 직접 입력")
        
    fetch_years = st.number_input("조회 기간 (년)", min_value=1, max_value=20, value=3)
    
    want_extra = st.checkbox("나라장터 기초금액·낙찰하한율 추가 조회", value=True)
    
    if st.button("🚀 낙찰 결과 조회 시작", type="primary"):
        key = clean_key(service_key)
        if not key:
            st.error("서비스키를 입력해주세요.")
        elif not selected_cd:
            st.error("기관 코드를 확인해주세요.")
        else:
            with st.spinner("조회 및 상세 분석 중입니다..."):
                try:
                    all_rows = []
                    wins = month_windows(fetch_years * 12)
                    for bgn, end in wins:
                        if sys_type == "g2b":
                            res = g2b_fetch_month(key, bgn, end, "2", code=selected_cd)
                            all_rows.extend([g2b_normalize(r) for r in res])
                        else:
                            res = d2b_fetch_month(key, bgn, end, code=selected_cd)
                            all_rows.extend([d2b_normalize(r) for r in res])
                    
                    # 나라장터인 경우 기초금액/낙찰하한율 추가 조회
                    if sys_type == "g2b" and want_extra:
                        progress_bar = st.progress(0)
                        total_items = len(all_rows)
                        for idx, r in enumerate(all_rows):
                            try:
                                extra = g2b_fetch_notice_extra(key, r["공고번호"], r["차수"])
                                if extra.get("bssamt"):
                                    r["기초금액"] = extra["bssamt"]
                                if extra.get("lwlt") is not None:
                                    r["낙찰하한율(%)"] = extra["lwlt"]
                                if extra.get("mthd"):
                                    r["낙찰방법"] = extra["mthd"]
                            except Exception:
                                pass
                            if total_items > 0:
                                progress_bar.progress((idx + 1) / total_items)
                        progress_bar.empty()

                    for r in all_rows:
                        calc_gap(r)
                        
                    st.session_state.rows = all_rows
                    st.success(f"조회 완료! 총 {len(all_rows):,}건")
                except Exception as e:
                    st.error(str(e))

with col2:
    st.subheader("📈 요약 및 엑셀 다운로드")
    rows = st.session_state.rows
    if rows:
        win_count = sum(1 for r in rows if r["결과"] == "낙찰")
        total_amt = sum(r["낙찰금액"] or 0 for r in rows)
        st.metric("총 조회 건수", f"{len(rows):,}건 (낙찰 {win_count:,}건)")
        st.metric("낙찰금액 합계", f"{total_amt:,}원")
        
        excel_data = generate_excel(rows)
        st.download_button(
            label="📥 엑셀 파일(.xlsx) 다운로드",
            data=excel_data,
            file_name=f"낙찰결과_{datetime.now():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("조회된 결과가 없습니다.")

if st.session_state.rows:
    st.divider()
    st.subheader("📋 결과 목록")
    filter_kw = st.text_input("🔍 공고명 검색", placeholder="검색할 공고명 입력")
    display_rows = st.session_state.rows
    if filter_kw:
        display_rows = [r for r in display_rows if filter_kw.lower() in r["공고명"].lower()]
    
    # 표에 보여줄 주요 항목 정돈
    view_data = []
    for r in display_rows:
        view_data.append({
            "개찰일시": r.get("개찰일시"),
            "공고번호": f"{r.get('공고번호')}-{r.get('차수')}" if r.get('차수') else r.get('공고번호'),
            "공고명": r.get("공고명"),
            "결과": r.get("결과"),
            "낙찰업체": r.get("낙찰업체"),
            "기초금액": fmt_amt(r.get("기초금액")),
            "낙찰하한율(%)": fmt_rate(r.get("낙찰하한율(%)")),
            "낙찰금액": fmt_amt(r.get("낙찰금액")),
            "낙찰률(%)": fmt_rate(r.get("낙찰률(%)")),
            "GAP(%p)": fmt_rate(r.get("GAP(%p)")),
            "낙찰방법": r.get("낙찰방법"),
        })
    st.dataframe(view_data, use_container_width=True)
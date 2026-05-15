"""아파트 실거래가 조회 — 국토부 공공데이터 (Streamlit 앱)

API 키 우선순위:
  1) Streamlit Secrets  (st.secrets["MOLIT_API_KEY"])  ← Streamlit Cloud 배포 시
  2) 환경변수           (os.environ["MOLIT_API_KEY"])
  3) 로컬 .env          (C:\\Users\\ajfxo\\.claude\\.env)             ← 지호님 로컬 PC
"""
import os
import time
import datetime as dt
from pathlib import Path
from collections import defaultdict
from xml.etree import ElementTree as ET

import requests
import pandas as pd
import streamlit as st

from lawd_codes import LAWD, list_sido, list_sigungu, get_code
from exporters import to_excel_bytes, to_html_bytes

API_URL = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTrade/getRTMSDataSvcAptTrade"
LOCAL_ENV_PATH = Path(r"C:\Users\ajfxo\.claude\.env")


# === API 키 로드 ===
@st.cache_resource
def load_key() -> str:
    try:
        key = st.secrets.get("MOLIT_API_KEY")
        if key:
            return str(key).strip()
    except Exception:
        pass

    key = os.environ.get("MOLIT_API_KEY")
    if key:
        return key.strip()

    if LOCAL_ENV_PATH.exists():
        for line in LOCAL_ENV_PATH.read_text(encoding="utf-8").splitlines():
            if line.startswith("MOLIT_API_KEY="):
                return line.split("=", 1)[1].strip().strip('"')

    st.error(
        "🔑 MOLIT_API_KEY가 설정되지 않았습니다.\n\n"
        "- 로컬 실행: `C:\\Users\\ajfxo\\.claude\\.env` 에 `MOLIT_API_KEY=...` 추가\n"
        "- Streamlit Cloud: 앱 설정 → Secrets 에 `MOLIT_API_KEY=\"...\"` 추가"
    )
    st.stop()


# === API 호출 (캐싱: 같은 시군구·월은 24시간 재호출 안 함) ===
@st.cache_data(ttl=60 * 60 * 24, show_spinner=False)
def fetch_month(lawd_cd: str, deal_ymd: str, _key: str) -> list[dict]:
    rows, page = [], 1
    while True:
        params = {"serviceKey": _key, "LAWD_CD": lawd_cd, "DEAL_YMD": deal_ymd,
                  "pageNo": page, "numOfRows": 1000}
        r = requests.get(API_URL, params=params, timeout=30)
        if r.status_code != 200:
            raise RuntimeError(f"API 오류 status={r.status_code}: {r.text[:200]}")
        root = ET.fromstring(r.text)
        result_code = root.findtext(".//resultCode")
        result_msg = (root.findtext(".//resultMsg") or "").strip().upper()
        if result_code and result_code not in ("00", "0", "000") and result_msg != "OK":
            raise RuntimeError(f"API 코드={result_code} 메시지={result_msg}")
        items = root.findall(".//item")
        if not items:
            break
        for it in items:
            rows.append({c.tag: (c.text or "").strip() for c in it})
        total = root.findtext(".//totalCount")
        if total and len(rows) >= int(total):
            break
        page += 1
        time.sleep(0.1)
    return rows


# === 데이터 가공 ===
def to_int_won(s: str) -> int:
    try:
        return int(s.replace(",", "").strip())
    except Exception:
        return 0


def round_pyeong(area_m2: float) -> float:
    return round((area_m2 / 3.305785) * 2) / 2


def build_raw_df(raw_rows: list[dict], sigungu_name: str) -> pd.DataFrame:
    out = []
    for r in raw_rows:
        try:
            area = float(r.get("excluUseAr", "0") or 0)
        except Exception:
            area = 0.0
        price = to_int_won(r.get("dealAmount", "0"))
        pyeong_exact = area / 3.305785 if area else 0
        ppp = int(price / pyeong_exact) if pyeong_exact else 0
        y = r.get("dealYear", "").zfill(4)
        m = r.get("dealMonth", "").zfill(2)
        d = r.get("dealDay", "").zfill(2)
        cdealDay = r.get("cdealDay", "").strip()
        out.append({
            "거래일자": f"{y}-{m}-{d}",
            "시군구": sigungu_name,
            "법정동": r.get("umdNm", ""),
            "단지명": r.get("aptNm", ""),
            "전용면적(㎡)": round(area, 2),
            "평형(평)": round(pyeong_exact, 1),
            "거래금액(만원)": price,
            "거래금액(억)": round(price / 10000, 2),
            "평단가(만원/평)": ppp,
            "층": r.get("floor", ""),
            "건축년도": r.get("buildYear", ""),
            "지번": r.get("jibun", ""),
            "거래종류": r.get("dealingGbn", ""),
            "해제여부": f"해제({cdealDay})" if cdealDay else "",
            "_평형군집": round_pyeong(area) if area else 0,
        })
    df = pd.DataFrame(out)
    if df.empty:
        return df
    return df.sort_values("거래일자", ascending=False).reset_index(drop=True)


def build_summary_df(raw_df: pd.DataFrame, include_canceled: bool) -> pd.DataFrame:
    if raw_df.empty:
        return pd.DataFrame()
    df = raw_df if include_canceled else raw_df[raw_df["해제여부"] == ""]
    grp = defaultdict(list)
    for _, row in df.iterrows():
        key = (row["시군구"], row["법정동"], row["단지명"], row["_평형군집"])
        grp[key].append(row)
    rows = []
    for (sgg, umd, apt, pyeong), items in grp.items():
        prices = [r["거래금액(억)"] for r in items]
        ppps = [r["평단가(만원/평)"] for r in items if r["평단가(만원/평)"]]
        floors = [int(r["층"]) for r in items if str(r["층"]).lstrip("-").isdigit()]
        latest = max(r["거래일자"] for r in items)
        rows.append({
            "시군구": sgg, "법정동": umd, "단지명": apt, "평형(평)": pyeong,
            "거래건수": len(items),
            "평균 거래가(억)": round(sum(prices) / len(prices), 2),
            "최고 거래가(억)": max(prices),
            "최저 거래가(억)": min(prices),
            "평균 평단가(만원/평)": int(sum(ppps) / len(ppps)) if ppps else 0,
            "최근 거래일자": latest,
            "비고": f"층 {min(floors)}~{max(floors)}" if floors else "",
        })
    return pd.DataFrame(rows).sort_values(
        ["시군구", "법정동", "단지명", "평형(평)"]
    ).reset_index(drop=True)


# === 헬퍼: 최근 24개월 YYYYMM 리스트 ===
def recent_months(n: int = 24) -> list[str]:
    today = dt.date.today().replace(day=1)
    out = []
    for i in range(n):
        y = today.year
        m = today.month - i
        while m <= 0:
            m += 12
            y -= 1
        out.append(f"{y:04d}{m:02d}")
    return out


def ymd_range(start_ymd: str, end_ymd: str) -> list[str]:
    """start_ymd ≤ x ≤ end_ymd 의 YYYYMM 리스트 (오름차순)"""
    sy, sm = int(start_ymd[:4]), int(start_ymd[4:])
    ey, em = int(end_ymd[:4]), int(end_ymd[4:])
    out = []
    y, m = sy, sm
    while (y, m) <= (ey, em):
        out.append(f"{y:04d}{m:02d}")
        m += 1
        if m > 12:
            m = 1
            y += 1
    return out


# === UI ===
st.set_page_config(page_title="아파트 실거래가 조회", page_icon="🏢", layout="wide")
st.title("🏢 아파트 실거래가 조회")
st.caption("국토교통부 공공데이터포털 · RTMSDataSvcAptTrade")

key = load_key()
months = recent_months(24)

# 사이드바
with st.sidebar:
    st.header("📍 지역")
    sido = st.selectbox("시도", list_sido(), index=1)  # 기본: 경기도
    multi = st.checkbox("시군구 다중선택", value=False)
    sigungu_opts = list_sigungu(sido)
    if multi:
        default_sgg = ["남양주시"] if "남양주시" in sigungu_opts else sigungu_opts[:1]
        sigungus = st.multiselect("시군구", sigungu_opts, default=default_sgg)
    else:
        default_idx = sigungu_opts.index("남양주시") if "남양주시" in sigungu_opts else 0
        sigungus = [st.selectbox("시군구", sigungu_opts, index=default_idx)]

    st.header("📅 기간")
    start_ymd = st.selectbox("시작월", months, index=0,
                             format_func=lambda x: f"{x[:4]}-{x[4:]}")
    end_ymd = st.selectbox("종료월", months, index=0,
                           format_func=lambda x: f"{x[:4]}-{x[4:]}")

    st.header("🔍 키워드")
    keyword = st.text_input("단지명·법정동 부분일치", placeholder="예: 덕소, 래미안 …")

    st.header("📐 필터")
    pyeong_range = st.slider("평형(평)", 0, 100, (0, 100))
    price_range = st.slider("가격(억)", 0, 100, (0, 100))
    include_canceled = st.checkbox("해제 거래 포함", value=False)

    go = st.button("🔥 조회", use_container_width=True, type="primary")

    st.markdown("---")
    st.caption(
        "© 2026 [Jiho-Hwang90](https://github.com/Jiho-Hwang90/Project) "
        "· Built with Streamlit"
    )


# === 조회 동작 ===
if go:
    if not sigungus:
        st.warning("시군구를 1개 이상 선택해주세요.")
        st.stop()

    # 시작 ≤ 종료 보정
    if start_ymd > end_ymd:
        start_ymd, end_ymd = end_ymd, start_ymd

    months_to_fetch = ymd_range(start_ymd, end_ymd)
    total_jobs = len(sigungus) * len(months_to_fetch)

    progress = st.progress(0, text="조회 시작...")
    raw_rows_per_sigungu: list[tuple[str, list[dict]]] = []
    job = 0
    for sgg in sigungus:
        try:
            code = get_code(sido, sgg)
        except KeyError:
            st.error(f"코드 없음: {sido} > {sgg}")
            continue
        for ymd in months_to_fetch:
            job += 1
            progress.progress(job / total_jobs,
                              text=f"[{job}/{total_jobs}] {sgg} {ymd[:4]}-{ymd[4:]}")
            try:
                rows = fetch_month(code, ymd, key)
            except Exception as e:
                st.error(f"{sgg} {ymd} 호출 실패: {e}")
                continue
            raw_rows_per_sigungu.append((sgg, rows))
    progress.empty()

    # 합치기
    raw_dfs = []
    for sgg, rows in raw_rows_per_sigungu:
        if rows:
            raw_dfs.append(build_raw_df(rows, sgg))
    if not raw_dfs:
        st.warning("조회 결과가 없습니다.")
        st.stop()
    raw_df = pd.concat(raw_dfs, ignore_index=True)

    # 필터 적용
    if keyword.strip():
        kw = keyword.strip()
        mask = raw_df["단지명"].str.contains(kw, na=False) | raw_df["법정동"].str.contains(kw, na=False)
        raw_df = raw_df[mask]
    raw_df = raw_df[
        (raw_df["평형(평)"] >= pyeong_range[0]) & (raw_df["평형(평)"] <= pyeong_range[1]) &
        (raw_df["거래금액(억)"] >= price_range[0]) & (raw_df["거래금액(억)"] <= price_range[1])
    ]
    if not include_canceled:
        raw_df = raw_df[raw_df["해제여부"] == ""]
    raw_df = raw_df.reset_index(drop=True)

    if raw_df.empty:
        st.warning("필터 조건에 맞는 거래가 없습니다.")
        st.stop()

    # 요약 시트용 df
    summary_df = build_summary_df(raw_df, include_canceled=True)

    # 요약 카드
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("총 거래", f"{len(raw_df):,}건")
    c2.metric("평균 거래가", f"{raw_df['거래금액(억)'].mean():.2f}억")
    ppp_valid = raw_df[raw_df["평단가(만원/평)"] > 0]["평단가(만원/평)"]
    c3.metric("평균 평단가", f"{int(ppp_valid.mean()):,}만/평" if len(ppp_valid) else "-")
    c4.metric("단지·평형 군집", f"{len(summary_df)}개")

    # 표 (출력용 컬럼만)
    summary_show = summary_df[["시군구", "법정동", "단지명", "평형(평)", "거래건수",
                               "평균 거래가(억)", "최고 거래가(억)", "최저 거래가(억)",
                               "평균 평단가(만원/평)", "최근 거래일자", "비고"]]
    raw_show = raw_df.drop(columns=["_평형군집"])

    tab1, tab2 = st.tabs(["📋 단지별 요약", "📑 전체 거래내역"])
    with tab1:
        st.dataframe(summary_show, use_container_width=True, hide_index=True)
    with tab2:
        st.dataframe(raw_show, use_container_width=True, hide_index=True)

    # 다운로드
    sgg_label = "_".join(sigungus) if len(sigungus) <= 3 else f"{sigungus[0]}외{len(sigungus)-1}"
    period_label = f"{start_ymd}-{end_ymd}" if start_ymd != end_ymd else start_ymd
    base_name = f"{period_label}_{sgg_label}_실거래가"
    gen_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    title = f"{sgg_label} 실거래가 ({period_label})"
    meta = f"필터: 평형 {pyeong_range[0]}~{pyeong_range[1]}평 / 가격 {price_range[0]}~{price_range[1]}억" \
           + (f" / 키워드 '{keyword}'" if keyword.strip() else "") \
           + (" / 해제 포함" if include_canceled else "")

    excel_bytes = to_excel_bytes(summary_show, raw_show)
    html_bytes = to_html_bytes(summary_show, raw_show, title, meta, gen_at)

    d1, d2 = st.columns(2)
    d1.download_button("📥 Excel 다운로드", data=excel_bytes,
                       file_name=f"{base_name}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
    d2.download_button("🌐 HTML 다운로드", data=html_bytes,
                       file_name=f"{base_name}.html",
                       mime="text/html",
                       use_container_width=True)

else:
    st.info("👈 좌측 사이드바에서 지역·기간·필터를 설정한 뒤 **조회** 버튼을 눌러주세요.")
    with st.expander("ℹ️ 이 앱은 무엇을 하나요?"):
        st.markdown("""
- **국토교통부 공공데이터포털**의 아파트매매 실거래가를 시군구·월 단위로 조회합니다.
- 데이터는 **계약일 기준**입니다. 신고 의무가 30일이라 최근 1~2개월 데이터는 누적 중일 수 있습니다.
- 동일 조건 재조회 시 캐시된 결과를 사용 (24시간 유지) — 일일 API 한도 절약.
- 결과는 Excel(2시트) · HTML(단일 파일) 두 가지로 다운로드 가능합니다.
- API 키는 PC 내부 `.env` 파일에서만 읽고, 외부로 전송하지 않습니다.
        """)

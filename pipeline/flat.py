# -*- coding: utf-8 -*-
"""Phase 4: 플랫표 (통합_캠페인일자별).

검색광고 제외(Naver SA + Google cpc), DA 실집행만.
행 = (날짜 × 브랜드 × 광고그룹), 유형/구분/지면 자동분류. 상단 합계.
"""
import warnings
warnings.simplefilter("ignore")
import re
from total import (excel_weeknum2, _put, F_TITLE, F_COL, F_SUM,
                   FILL_COL, FILL_SUM, CENTER, LEFT, _div, SAT_COLOR, SUN_COLOR)

FLAT_COLS = ["주차", "날짜", "브랜드", "유형", "구분", "광고그룹", "지면",
             "노출수", "클릭수", "클릭률", "클릭당비용", "집행예산", "전환율",
             "거래수", "수익", "회원가입", "회원가입율", "ROAS", "객단가"]
# 지표컬럼 서식 (8열부터)
FLAT_FMT = ["#,##0", "#,##0", "0.00%", "#,##0", "#,##0", "0.00%",
            "#,##0", "#,##0", "#,##0", "0.00%", "#,##0.00", "#,##0"]


def _is_search(media, camp):
    if media == "Naver SA":
        return True
    if media == "Google" and "cpc" in str(camp).lower():
        return True
    return False


def _type(camp):
    return "노출형" if "_br_" in str(camp).lower() else "성과형"


def _gubun(adgroup):
    return "기획전" if re.search(r"(NEW|RE)_\d", str(adgroup)) else "일반"


def _jimyeon(media, adgroup):
    if media == "Meta":
        return "Meta ASC" if "ASC" in str(adgroup) else "Meta"
    return {"KKO": "KKO", "Criteo": "Criteo", "Naver": "Naver",
            "RTB": "RTB하우스", "Google": "Google"}.get(media, media)


def build_flat(uni):
    """플랫표 행 리스트 반환 (dict). DA 실집행만."""
    da = uni[~uni.apply(lambda r: _is_search(r["매체"], r["캠페인"]), axis=1)].copy()
    da = da[(da["노출수"] > 0) | (da["클릭수"] > 0)]     # 실집행
    g = da.groupby(["날짜", "날짜키", "브랜드", "매체", "광고그룹"], sort=True).agg(
        imp=("노출수", "sum"), clk=("클릭수", "sum"), cost=("광고비용", "sum"),
        cv=("GA구매", "sum"), rev=("GA구매수익", "sum"), mem=("회원가입수", "sum"),
        camp=("캠페인", "first")).reset_index()
    rows = []
    for _, r in g.iterrows():
        rows.append({
            "주차": excel_weeknum2(r["날짜"].date()), "날짜": r["날짜"], "브랜드": r["브랜드"],
            "유형": _type(r["camp"]), "구분": _gubun(r["광고그룹"]),
            "광고그룹": r["광고그룹"], "지면": _jimyeon(r["매체"], r["광고그룹"]),
            "노출수": r["imp"], "클릭수": r["clk"], "클릭률": _div(r["clk"], r["imp"]),
            "클릭당비용": _div(r["cost"], r["clk"]), "집행예산": r["cost"],
            "전환율": _div(r["cv"], r["clk"]), "거래수": r["cv"], "수익": r["rev"],
            "회원가입": r["mem"], "회원가입율": _div(r["mem"], r["clk"]),
            "ROAS": _div(r["rev"], r["cost"]), "객단가": _div(r["rev"], r["cv"]),
        })
    return rows


def write_flat(ws, uni, y, mth):
    rows = build_flat(uni)
    _put(ws, 2, 1, "통합_캠페인일자별 (자동)", font=F_TITLE)
    hr = 3
    for i, h in enumerate(FLAT_COLS):
        _put(ws, hr, 1 + i, h, font=F_COL, fill=FILL_COL, align=CENTER)
    # 합계 (헤더 아래)
    tot = {k: sum(r[k] for r in rows) for k in ["노출수", "클릭수", "집행예산", "거래수", "수익", "회원가입"]}
    sr = hr + 1
    _put(ws, sr, 1, "합계", font=F_SUM, fill=FILL_SUM)
    agg = {"노출수": tot["노출수"], "클릭수": tot["클릭수"],
           "클릭률": _div(tot["클릭수"], tot["노출수"]), "클릭당비용": _div(tot["집행예산"], tot["클릭수"]),
           "집행예산": tot["집행예산"], "전환율": _div(tot["거래수"], tot["클릭수"]),
           "거래수": tot["거래수"], "수익": tot["수익"], "회원가입": tot["회원가입"],
           "회원가입율": _div(tot["회원가입"], tot["클릭수"]), "ROAS": _div(tot["수익"], tot["집행예산"]),
           "객단가": _div(tot["수익"], tot["거래수"])}
    for i, k in enumerate(FLAT_COLS[7:]):
        _put(ws, sr, 8 + i, agg[k], FLAT_FMT[i], font=F_SUM, fill=FILL_SUM)
    # 데이터
    r0 = sr + 1
    for j, rec in enumerate(rows):
        row = r0 + j
        wd = rec["날짜"].weekday()
        col = SAT_COLOR if wd == 5 else SUN_COLOR if wd == 6 else None
        _put(ws, row, 1, rec["주차"], align=CENTER)
        _put(ws, row, 2, rec["날짜"], "yyyy-mm-dd", align=CENTER, color=col)
        for i, k in enumerate(["브랜드", "유형", "구분", "광고그룹", "지면"]):
            _put(ws, row, 3 + i, rec[k], align=LEFT)
        for i, k in enumerate(FLAT_COLS[7:]):
            _put(ws, row, 8 + i, rec[k], FLAT_FMT[i])
    # 열폭
    for c, w in [(1, 5), (2, 12), (3, 7), (4, 8), (5, 8), (6, 26), (7, 11)]:
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = w
    for c in range(8, 20):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 11

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
            "#,##0", "#,##0", "#,##0", "0.00%", "#,##0%", "#,##0"]


def _is_search(media, camp):
    if media == "Naver SA":
        return True
    if media == "Google" and "cpc" in str(camp).lower():
        return True
    return False


def _type(camp):
    return "노출형" if "_br_" in str(camp).lower() else "성과형"


def _gubun(adgroup):
    # 기획전 = NEW/RE 뒤 기획전코드(숫자) 또는 기간(MMDD-MMDD) 포함. 그 외 일반.
    s = str(adgroup)
    if re.search(r"(NEW|RE)_\d", s) or re.search(r"\d{4}-\d{4}", s):
        return "기획전"
    return "일반"


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
    """A열 비움(B부터). 합계=SUBTOTAL(필터반영), 헤더 AutoFilter, 틀고정."""
    rows = build_flat(uni)
    C0 = 2                                  # 데이터 시작열 = B
    _put(ws, 2, C0, "통합_캠페인일자별", font=F_TITLE)
    sum_r, hdr_r, r0 = 3, 4, 5              # 합계행 / 헤더행 / 데이터시작
    last = r0 + len(rows) - 1 if rows else r0

    # 헤더 (row4)
    for i, h in enumerate(FLAT_COLS):
        _put(ws, hdr_r, C0 + i, h, font=F_COL, fill=FILL_COL, align=CENTER)

    # 합계 (row3, SUBTOTAL — 필터 시 자동 반영). 지표는 col C0+7 부터
    def L(offset):
        return ws.cell(row=1, column=C0 + offset).column_letter
    # B3(합계 라벨) 셀은 비움(색·텍스트 없음)
    rng = f"{r0}:{last}"
    # 합계 컬럼 오프셋: 노출7 클릭8 클릭률9 클릭당비용10 집행예산11 전환율12 거래수13 수익14 회원가입15 회원가입율16 ROAS17 객단가18
    sub = {7: None, 8: None, 11: None, 13: None, 14: None, 15: None}   # SUBTOTAL 대상(합)
    for off in sub:
        c = L(off)
        _put(ws, sum_r, C0 + off, f"=SUBTOTAL(109,{c}{r0}:{c}{last})", FLAT_FMT[off - 7],
             font=F_SUM, fill=FILL_SUM)
    # 비율 = 합계행 셀 참조
    ratio = {
        9:  f"=IFERROR({L(8)}{sum_r}/{L(7)}{sum_r},0)",    # 클릭률=클릭/노출
        10: f"=IFERROR({L(11)}{sum_r}/{L(8)}{sum_r},0)",   # 클릭당비용=집행/클릭
        12: f"=IFERROR({L(13)}{sum_r}/{L(8)}{sum_r},0)",   # 전환율=거래/클릭
        16: f"=IFERROR({L(15)}{sum_r}/{L(8)}{sum_r},0)",   # 회원가입율=회원/클릭
        17: f"=IFERROR({L(14)}{sum_r}/{L(11)}{sum_r},0)",  # ROAS=수익/집행
        18: f"=IFERROR({L(14)}{sum_r}/{L(13)}{sum_r},0)",  # 객단가=수익/거래
    }
    for off, formula in ratio.items():
        _put(ws, sum_r, C0 + off, formula, FLAT_FMT[off - 7], font=F_SUM, fill=FILL_SUM)

    # 데이터 (row5~)
    for j, rec in enumerate(rows):
        row = r0 + j
        wd = rec["날짜"].weekday()
        col = SAT_COLOR if wd == 5 else SUN_COLOR if wd == 6 else None
        _put(ws, row, C0, rec["주차"], align=CENTER)
        _put(ws, row, C0 + 1, rec["날짜"], "yyyy-mm-dd", align=CENTER, color=col)
        for i, k in enumerate(["브랜드", "유형", "구분", "광고그룹", "지면"]):
            _put(ws, row, C0 + 2 + i, rec[k], align=LEFT)
        for i, k in enumerate(FLAT_COLS[7:]):
            _put(ws, row, C0 + 7 + i, rec[k], FLAT_FMT[i])

    # AutoFilter(헤더~데이터) + 틀고정(합계·헤더 항상 보이게)
    ws.auto_filter.ref = f"{L(0)}{hdr_r}:{L(len(FLAT_COLS) - 1)}{last}"
    ws.freeze_panes = ws.cell(row=r0, column=1).coordinate   # A5: 위 4행 고정

    # 열폭 (A 비움)
    ws.column_dimensions["A"].width = 3
    widths = [(0, 5), (1, 12), (2, 7), (3, 8), (4, 8), (5, 26), (6, 11)]
    for off, w in widths:
        ws.column_dimensions[L(off)].width = w
    for off in range(7, len(FLAT_COLS)):
        ws.column_dimensions[L(off)].width = 11

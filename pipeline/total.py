# -*- coding: utf-8 -*-
"""Phase 2: 브랜드 Total 대시보드 (MI/IT/EBM).

통합에서 브랜드별 슬라이스 → 6개 섹션 값 산출 후 시트 작성:
 [매체 예산 집행율] [매체 총 누적] [디바이스별=구분롤업] [요일별 평균] [주간] [일자별 성과]
"""
import warnings
warnings.simplefilter("ignore")
import calendar
from datetime import date, timedelta
import pandas as pd

BRAND_TITLE = {"MI": "미샤", "IT": "잇미샤", "EBM": "E.B.M"}

# (구분, 라벨, 매체, 캠페인 부분문자열, 월예산)   부분문자열 ""=매체 전체
MEDIA_ROWS = [
    ("SA",        "네이버 브랜드검색", "Naver SA", "bsa",        11_330_000),
    ("SA",        "네이버 키워드검색", "Naver SA", "cpc",         3_137_000),
    ("SA",        "네이버 쇼핑검색",   "Naver SA", "shopping",    2_000_000),
    ("SA",        "네이버 엠버서더형", "Naver SA", "Ambassador",  2_442_000),
    ("SA",        "구글 키워드검색",   "Google",   "cpc",         1_500_000),
    ("DA(성과형)", "카카오 네이티브",   "KKO",      "ntv",           940_000),
    ("DA(성과형)", "카카오 비즈보드",   "KKO",      "biz",           600_000),
    ("DA(성과형)", "카카오 카탈로그",   "KKO",      "_ca",         3_000_000),
    ("DA(성과형)", "구글 쇼핑",        "Google",   "pmax",        9_191_000),
    ("DA(성과형)", "크리테오",         "Criteo",   "",           17_000_000),
    ("DA(성과형)", "인스타그램 ",      "Meta",     "pf",          8_160_000),
    ("DA(성과형)", "RTB",             "RTB",      "",            2_700_000),
    ("DA(성과형)", "구글 GDN",         "Google",   "gdn",                 0),
    ("DA(노출형)", "인스타그램 ",      "Meta",     "br",          2_200_000),
    ("DA(노출형)", "구글 YouTube",     "Google",   "youtube",             0),
]

CUM_HDR = ["노출수", "클릭수", "클릭률", "클릭당비용", "집행예산", "전환수", "매출",
           "회원가입", "전환율", "전환당비용", "회원가입율", "ROAS", "객단가"]
WD_KR = ["월", "화", "수", "목", "금", "토", "일"]


def _div(a, b):
    return a / b if b else 0.0


def _slice(df, media, pattern):
    m = df["매체"] == media
    if pattern:
        m &= df["캠페인"].str.contains(pattern, case=False, regex=False)
    return df[m]


def _metrics(sub):
    imp = sub["노출수"].sum(); clk = sub["클릭수"].sum(); cost = sub["광고비용"].sum()
    cv = sub["GA구매"].sum(); rev = sub["GA구매수익"].sum()
    mem = sub["회원가입수"].sum()
    return {
        "노출수": imp, "클릭수": clk, "클릭률": _div(clk, imp), "클릭당비용": _div(cost, clk),
        "집행예산": cost, "전환수": cv, "매출": rev, "회원가입": mem,
        "전환율": _div(cv, clk), "전환당비용": _div(cost, cv), "회원가입율": _div(mem, clk),
        "ROAS": _div(rev, cost), "객단가": _div(rev, cv),
    }


def _metrics_from_sums(imp, clk, cost, cv, rev, mem):
    return {
        "노출수": imp, "클릭수": clk, "클릭률": _div(clk, imp), "클릭당비용": _div(cost, clk),
        "집행예산": cost, "전환수": cv, "매출": rev, "회원가입": mem,
        "전환율": _div(cv, clk), "전환당비용": _div(cost, cv), "회원가입율": _div(mem, clk),
        "ROAS": _div(rev, cost), "객단가": _div(rev, cv),
    }


def excel_weeknum2(d):
    jan1 = date(d.year, 1, 1)
    start = jan1 - timedelta(days=jan1.weekday())     # Jan1 주의 월요일
    return (d - start).days // 7 + 1


def week_in_month(d):
    first = d.replace(day=1)
    return excel_weeknum2(d) - excel_weeknum2(first) + 1


def daily_frame(df_brand, y, mth):
    """1일~월말 전체 일자 성과 (빠진 날은 0)."""
    ndays = calendar.monthrange(y, mth)[1]
    g = df_brand.groupby("날짜키").agg(
        imp=("노출수", "sum"), clk=("클릭수", "sum"), cost=("광고비용", "sum"),
        cv=("GA구매", "sum"), rev=("GA구매수익", "sum"), mem=("회원가입수", "sum")).to_dict("index")
    recs = []
    for day in range(1, ndays + 1):
        d = date(y, mth, day)
        dk = d.strftime("%Y%m%d")
        s = g.get(dk, {"imp": 0, "clk": 0, "cost": 0, "cv": 0, "rev": 0, "mem": 0})
        m = _metrics_from_sums(s["imp"], s["clk"], s["cost"], s["cv"], s["rev"], s["mem"])
        recs.append({"주차": week_in_month(d), "요일": WD_KR[d.weekday()], "날짜": d,
                     "wd": d.weekday(), **m})
    return pd.DataFrame(recs)


if __name__ == "__main__":
    from build import build_unified
    uni = build_unified()
    y, mth = 2026, 7
    for b in ["MI", "IT", "EBM"]:
        sub = uni[uni["브랜드"] == b]
        tot = _metrics(sub)
        df = daily_frame(sub, y, mth)
        wk = df.groupby("주차")[["집행예산", "전환수", "매출"]].sum()
        print(f"\n[{b}] 합계 광고비={tot['집행예산']:,.0f} 전환={tot['전환수']:.0f} 매출={tot['매출']:,.0f}")
        print("  주간 광고비:", {int(k): round(v) for k, v in wk["집행예산"].items()})

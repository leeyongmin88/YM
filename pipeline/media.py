# -*- coding: utf-8 -*-
"""Phase 3: 매체별 상세 리포트 (브랜드 × 매체 시트).

단일캠페인형(구글SA·피맥스·RTB): ■주간현황 → ■요일별 평균 → ■일자별 성과
지표 순서·서식은 참고파일 기준.
"""
import warnings
warnings.simplefilter("ignore")
from total import (daily_frame, weekday_avg, _metrics, _metrics_from_sums, _div,
                   F_TITLE, F_SEC, F_COL, F_SUM, FILL_SEC, FILL_COL, FILL_SUM,
                   SAT_COLOR, SUN_COLOR, CENTER, LEFT, _put, BRAND_TITLE, WD_KR)

BRAND_LOWER = {"MI": "mi", "IT": "it", "EBM": "ebm"}

# 매체 리포트 지표: (표시라벨, 지표키, 서식)  — 참고파일 순서(회원가입율→ROAS)
MEDIA_COLS = [
    ("노출수", "노출수", "#,##0"), ("클릭수", "클릭수", "#,##0"),
    ("클릭률", "클릭률", "0.00%"), ("클릭당비용", "클릭당비용", "#,##0"),
    ("광고비", "집행예산", "#,##0"), ("전환수", "전환수", "#,##0"),
    ("매출", "매출", "#,##0"), ("회원가입", "회원가입", "#,##0"),
    ("세션수", "세션수", "#,##0"), ("전환율", "전환율", "0.00%"),
    ("전환당비용", "전환당비용", "#,##0"), ("회원가입율", "회원가입율", "0.00%"),
    ("ROAS", "ROAS", "#,##0.00"), ("객단가", "객단가", "#,##0"),
]


def _filter(uni, brand, media, pattern):
    m = (uni["브랜드"] == brand) & (uni["매체"] == media)
    if pattern:
        m &= uni["캠페인"].str.contains(pattern, case=False, regex=False)
    return uni[m]


def week_periods(daily):
    """주차별 기간 라벨 (mm/dd~mm/dd). daily_frame은 전월일 포함."""
    out = {}
    for wk in range(1, 6):
        sub = daily[daily["주차"] == wk]
        if len(sub):
            a, b = sub["날짜"].min(), sub["날짜"].max()
            out[wk] = f"{a.month:02d}/{a.day:02d}~{b.month:02d}/{b.day:02d}"
        else:
            out[wk] = ""
    return out


def _metric_row(ws, r, m, start_col=4):
    for i, (_, key, fmt) in enumerate(MEDIA_COLS):
        _put(ws, r, start_col + i, m.get(key, 0), fmt)


def _hdr(ws, r, first, second=None):
    _put(ws, r, 2, first, font=F_COL, fill=FILL_COL, align=CENTER)
    if second:
        _put(ws, r, 3, second, font=F_COL, fill=FILL_COL, align=CENTER)
    for i, (label, _, _) in enumerate(MEDIA_COLS):
        _put(ws, r, 4 + i, label, font=F_COL, fill=FILL_COL, align=CENTER)


def write_media_single(ws, brand, title, media_disp, camp_disp, df_f, y, mth):
    daily = daily_frame(df_f, y, mth)
    total = _metrics(df_f)
    periods = week_periods(daily)

    _put(ws, 2, 2, title, font=F_TITLE)
    _put(ws, 3, 2, "브랜드"); _put(ws, 3, 3, BRAND_LOWER[brand])
    _put(ws, 4, 2, "매체"); _put(ws, 4, 3, media_disp)
    _put(ws, 5, 2, "캠페인"); _put(ws, 5, 3, camp_disp)
    r = 7

    # ■ 주간현황
    _put(ws, r, 2, "■ 주간현황", font=F_SEC, fill=FILL_SEC); r += 1
    _hdr(ws, r, "주차", "기간"); r += 1
    for wk in range(1, 6):
        sub = daily[daily["주차"] == wk]
        m = _metrics_from_sums(sub["노출수"].sum(), sub["클릭수"].sum(), sub["집행예산"].sum(),
                               sub["전환수"].sum(), sub["매출"].sum(), sub["회원가입"].sum(),
                               sub["세션수"].sum())
        _put(ws, r, 2, f"{wk}주차", align=CENTER)
        _put(ws, r, 3, periods[wk], align=CENTER)
        _metric_row(ws, r, m); r += 1
    _put(ws, r, 2, "합계", font=F_SUM, fill=FILL_SUM)
    _metric_row(ws, r, total); r += 2

    # ■ 요일별 평균
    _put(ws, r, 2, "■ 요일별 평균", font=F_SEC, fill=FILL_SEC); r += 1
    _hdr(ws, r, "요일"); r += 1
    wrows, _, _, _ = weekday_avg(daily, y, mth)
    for wr in wrows:
        col = SAT_COLOR if wr["wd"] == 5 else SUN_COLOR if wr["wd"] == 6 else None
        _put(ws, r, 2, wr["요일"], align=CENTER, color=col)
        # weekday_avg keys: 노출,클릭,광고비,전환,매출... 다른 키명 → 매핑
        wm = {"노출수": wr["노출"], "클릭수": wr["클릭"], "클릭률": wr["클릭율"],
              "클릭당비용": wr["클릭비용"], "집행예산": wr["광고비"], "전환수": wr["전환"],
              "매출": wr["매출"], "전환당비용": wr["전환비용"], "ROAS": wr["ROAS"],
              "객단가": wr["객단가"], "회원가입": 0, "세션수": 0, "전환율": 0, "회원가입율": 0}
        _metric_row(ws, r, wm); r += 1
    r += 1

    # ■ 일자별 성과
    _put(ws, r, 2, "■ 일자별 성과", font=F_SEC, fill=FILL_SEC); r += 1
    _hdr(ws, r, "요일", "날짜"); r += 1
    for _, d in daily.iterrows():
        col = SAT_COLOR if d["wd"] == 5 else SUN_COLOR if d["wd"] == 6 else None
        _put(ws, r, 2, d["요일"], align=CENTER, color=col)
        _put(ws, r, 3, d["날짜"], "yyyy-mm-dd", align=CENTER, color=col)
        _metric_row(ws, r, d); r += 1
    _put(ws, r, 2, "합계", font=F_SUM, fill=FILL_SUM)
    _metric_row(ws, r, total)

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 13
    for c in range(4, 4 + len(MEDIA_COLS)):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 12


def write_media_multi(ws, brand, title, media_disp, group_col, df_f, y, mth):
    """다중유형 리포트: ■누적요약(유형별) → ■주간현황 → ■전체 일별 성과."""
    daily = daily_frame(df_f, y, mth)
    total = _metrics(df_f)
    periods = week_periods(daily)

    _put(ws, 2, 2, title, font=F_TITLE)
    _put(ws, 3, 2, "브랜드"); _put(ws, 3, 3, BRAND_LOWER[brand])
    _put(ws, 4, 2, "매체"); _put(ws, 4, 3, media_disp)
    r = 6

    # ■ 누적 요약 (유형별)
    _put(ws, r, 2, "■ 누적 요약 (유형별)", font=F_SEC, fill=FILL_SEC); r += 1
    hdr_label = {"캠페인": "광고유형", "광고그룹": "광고그룹", "상품유형": "상품"}.get(group_col, "구분")
    _put(ws, r, 2, hdr_label, font=F_COL, fill=FILL_COL, align=CENTER)
    _put(ws, r, 3, "캠페인", font=F_COL, fill=FILL_COL, align=CENTER)
    for i, (label, _, _) in enumerate(MEDIA_COLS):
        _put(ws, r, 4 + i, label, font=F_COL, fill=FILL_COL, align=CENTER)
    r += 1
    # 유형별 집계 (광고비 내림차순)
    grp = []
    for gval, sub in df_f.groupby(group_col):
        grp.append((str(gval), sub))
    grp.sort(key=lambda x: -x[1]["광고비용"].sum())
    for gval, sub in grp:
        m = _metrics(sub)
        camp = sub["캠페인"].iloc[0] if len(sub) else ""
        _put(ws, r, 2, gval[:24], align=LEFT)
        _put(ws, r, 3, camp[:22], align=LEFT)
        _metric_row(ws, r, m); r += 1
    _put(ws, r, 2, "TOTAL", font=F_SUM, fill=FILL_SUM)
    _metric_row(ws, r, total); r += 2

    # ■ 주간현황
    _put(ws, r, 2, "■ 주간현황", font=F_SEC, fill=FILL_SEC); r += 1
    _hdr(ws, r, "주차", "기간"); r += 1
    for wk in range(1, 6):
        sub = daily[daily["주차"] == wk]
        m = _metrics_from_sums(sub["노출수"].sum(), sub["클릭수"].sum(), sub["집행예산"].sum(),
                               sub["전환수"].sum(), sub["매출"].sum(), sub["회원가입"].sum(),
                               sub["세션수"].sum())
        _put(ws, r, 2, f"{wk}주차", align=CENTER)
        _put(ws, r, 3, periods[wk], align=CENTER)
        _metric_row(ws, r, m); r += 1
    _put(ws, r, 2, "합계", font=F_SUM, fill=FILL_SUM)
    _metric_row(ws, r, total); r += 2

    # ■ 전체 일별 성과
    _put(ws, r, 2, "■ 전체 일별 성과", font=F_SEC, fill=FILL_SEC); r += 1
    _hdr(ws, r, "요일", "날짜"); r += 1
    for _, d in daily.iterrows():
        col = SAT_COLOR if d["wd"] == 5 else SUN_COLOR if d["wd"] == 6 else None
        _put(ws, r, 2, d["요일"], align=CENTER, color=col)
        _put(ws, r, 3, d["날짜"], "yyyy-mm-dd", align=CENTER, color=col)
        _metric_row(ws, r, d); r += 1
    _put(ws, r, 2, "합계", font=F_SUM, fill=FILL_SUM)
    _metric_row(ws, r, total)

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 18
    for c in range(4, 4 + len(MEDIA_COLS)):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 12


# 다중유형 시트: (접미사, 제목, 매체표시, 매체, 패턴, 그룹컬럼, 대상브랜드)
MULTI_SHEETS = [
    ("크리테오", "{T} 크리테오 리포트", "criteo", "Criteo", "", "캠페인", ["MI", "IT", "EBM"]),
    ("K디스", "{T} 카카오 디스플레이 리포트", "kakao", "KKO", "", "캠페인", ["MI", "IT", "EBM"]),
    ("메타_성과형", "{T} 메타 성과형(pf) 리포트", "meta", "Meta", "pf", "광고그룹", ["MI", "IT", "EBM"]),
    ("메타_브랜딩형", "{T} 메타 브랜딩형(br) 리포트", "meta", "Meta", "br", "광고그룹", ["MI", "IT", "EBM"]),
    ("N디스", "{T} 네이버 디스플레이 리포트", "naver", "Naver", "", "캠페인", ["IT"]),
]


# 단일캠페인형 시트 정의: (접미사, 제목템플릿, 매체표시, 캠페인표시템플릿, 매체, 패턴, 대상브랜드)
SINGLE_SHEETS = [
    ("구글SA", "{T} 구글 SA 리포트", "google", "GGL_{B}_SA_pf_cpc", "Google", "cpc", ["MI", "IT", "EBM"]),
    ("피맥스_리포트", "{T} 피맥스 리포트", "google", "GGL_{B}_SA_pf_pmax", "Google", "pmax", ["MI", "IT", "EBM"]),
    ("RTB", "{T} RTB 리포트", "rtbhouse", "{B}_pf", "RTB", "", ["MI", "IT"]),
]


def add_media_sheets(book, uni, y, mth):
    for suffix, title_t, mdisp, camp_t, media, pat, brands in SINGLE_SHEETS:
        for b in brands:
            df_f = _filter(uni, b, media, pat)
            ws = book.create_sheet(f"{b}_{suffix}")
            write_media_single(ws, b, title_t.format(T=BRAND_TITLE[b]), mdisp,
                               camp_t.format(B=b), df_f, y, mth)
    for suffix, title_t, mdisp, media, pat, gcol, brands in MULTI_SHEETS:
        for b in brands:
            df_f = _filter(uni, b, media, pat)
            ws = book.create_sheet(f"{b}_{suffix}")
            write_media_multi(ws, b, title_t.format(T=BRAND_TITLE[b]), mdisp, gcol, df_f, y, mth)
    # N검색 (Naver SA, 상품유형별) — 세로형식
    for b in ["MI", "IT", "EBM"]:
        df_f = _filter(uni, b, "Naver SA", "").copy()
        df_f["상품유형"] = df_f["캠페인"].map(nsearch_type)
        ws = book.create_sheet(f"{b}_N검색")
        write_media_multi(ws, b, f"{BRAND_TITLE[b]} N검색 리포트", "naver", "상품유형", df_f, y, mth)


NSEARCH_TYPES = [("브랜드검색", "bsa"), ("파워링크", "cpc"),
                 ("쇼핑검색", "shopping"), ("엠버서더", "Ambassador"), ("플레이스", "place")]


def nsearch_type(camp):
    c = str(camp).lower()
    for label, pat in NSEARCH_TYPES:
        if pat.lower() in c:
            return label
    return "기타"


if __name__ == "__main__":
    from build import build_unified
    uni = build_unified()
    for b in ["MI", "IT", "EBM"]:
        df_f = _filter(uni, b, "Google", "cpc")
        t = _metrics(df_f)
        print(f"{b}_구글SA 합계: 노출={t['노출수']:,.0f} 클릭={t['클릭수']:,.0f} "
              f"광고비={t['집행예산']:,.0f} 전환={t['전환수']:.0f} 세션={t['세션수']:.0f}")

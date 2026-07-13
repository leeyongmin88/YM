# -*- coding: utf-8 -*-
"""Phase 5: 브랜드 종합 + 리포트 추가 요청.

- 브랜드 종합: 브랜드별 [매체 총 누적] 세로 스택.
- 리포트 추가 요청: 일별 광고비 매트릭스 (브랜드×매체세부 행 × 날짜 열 + 월누계).
"""
import warnings
warnings.simplefilter("ignore")
import calendar
from datetime import date
from openpyxl.styles import Alignment
from total import (MEDIA_ROWS, _slice, _metrics, _put, F_TITLE, F_SEC, F_COL, F_SUM,
                   FILL_SEC, FILL_COL, FILL_SUM, CENTER, LEFT, BRAND_TITLE)

# 브랜드 종합 지표 (참고파일 순서)
BS_KEYS = ["노출수", "클릭수", "클릭률", "클릭당비용", "집행예산", "전환수", "매출",
           "회원가입", "전환율", "전환당비용", "회원가입율", "ROAS", "객단가"]
BS_FMT = ["#,##0", "#,##0", "0.00%", "#,##0", "#,##0", "#,##0", "#,##0",
          "#,##0", "0.00%", "#,##0", "0.00%", "#,##0.00", "#,##0"]


def write_brand_summary(ws, uni, y, mth):
    _put(ws, 2, 2, f"- {y}년 {mth}월 -", font=F_TITLE)
    _put(ws, 3, 2, "시선인터내셔널 Ad Report", font=F_TITLE)
    _put(ws, 5, 2, "[브랜드 별 매체 총 누적]", font=F_SEC, fill=FILL_SEC)
    hdr = ["브랜드", "유형", "매체", "진행율"] + BS_KEYS
    for i, h in enumerate(hdr):
        _put(ws, 6, 2 + i, h, font=F_COL, fill=FILL_COL, align=CENTER)
    r = 7
    for b in ["MI", "EBM", "IT"]:
        dfb = uni[uni["브랜드"] == b]
        first = True
        for gubun, label, media, pat, budget in MEDIA_ROWS:
            m = _metrics(_slice(dfb, media, pat))
            _put(ws, r, 2, b if first else None, align=CENTER, font=F_SUM if first else None)
            _put(ws, r, 3, "노출형" if gubun == "DA(노출형)" else "성과형", align=CENTER)
            _put(ws, r, 4, label, align=LEFT)
            _put(ws, r, 5, "-", align=CENTER)
            for i, k in enumerate(BS_KEYS):
                _put(ws, r, 6 + i, m[k], BS_FMT[i])
            first = False
            r += 1
        # 브랜드 소계
        mt = _metrics(dfb)
        _put(ws, r, 2, f"{b} 소계", font=F_SUM, fill=FILL_SUM)
        for i, k in enumerate(BS_KEYS):
            _put(ws, r, 6 + i, mt[k], BS_FMT[i], font=F_SUM, fill=FILL_SUM)
        r += 2
    ws.column_dimensions["A"].width = 3
    for c, w in [(2, 8), (3, 8), (4, 18), (5, 7)]:
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = w
    for c in range(6, 6 + len(BS_KEYS)):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 12


# 리포트 추가 요청: 매체 세부 16종 (6월F 순서). 모든 브랜드에 전 매체 행 표시.
SUBTYPES = [
    ("Google Pmax", "Google", "pmax"), ("Google Keyword", "Google", "cpc"),
    ("Naver Brand SA", "Naver SA", "bsa"), ("Naver Keyword SA", "Naver SA", "cpc"),
    ("Naver Shopping SA", "Naver SA", "shopping"), ("Naver Place SA", "Naver SA", "place"),
    ("Naver Ambassador", "Naver SA", "Ambassador"),
    ("Naver DA_Smart", "Naver", "smart"), ("Naver DA_ADVoost", "Naver", "advoost"),
    ("KAKAO Bizboard", "KKO", "biz"), ("KAKAO Native", "KKO", "ntv"),
    ("KAKAO Catalog", "KKO", "ca"),
    ("Criteo", "Criteo", ""), ("RTB House", "RTB", ""),
    ("Instagram_성과형", "Meta", "pf"), ("Instagram_노출형(br)", "Meta", "br"),
]


def write_report_request(ws, uni, y, mth):
    ndays = calendar.monthrange(y, mth)[1]
    _put(ws, 2, 2, "■ 광고비용 데일리 집행현황", font=F_TITLE)
    # (3행 공백) 헤더 (row 4): 브랜드, 매체, 날짜1..N, 월누계
    _put(ws, 4, 2, "브랜드", font=F_COL, fill=FILL_COL, align=CENTER)
    _put(ws, 4, 3, "매체", font=F_COL, fill=FILL_COL, align=CENTER)
    for d in range(1, ndays + 1):
        _put(ws, 4, 3 + d, date(y, mth, d), "mm-dd", font=F_COL, fill=FILL_COL, align=CENTER)
    _put(ws, 4, 4 + ndays, "월 누계", font=F_COL, fill=FILL_COL, align=CENTER)

    def cost_series(dfb, media, pat):
        d = dfb[dfb["매체"] == media]
        if pat:
            d = d[d["캠페인"].str.contains(pat, case=False, regex=False)]
        return d.groupby("날짜키")["광고비용"].sum().to_dict()

    r = 5
    center_v = Alignment(horizontal="center", vertical="center")
    for b in ["전체", "MI", "EBM", "IT"]:
        dfb = uni if b == "전체" else uni[uni["브랜드"] == b]
        start = r
        for label, media, pat in SUBTYPES:      # 모든 매체 행 표시(0이어도)
            series = cost_series(dfb, media, pat)
            _put(ws, r, 3, label, align=LEFT)
            total = 0.0
            for dd in range(1, ndays + 1):
                dk = date(y, mth, dd).strftime("%Y%m%d")
                v = series.get(dk, 0)
                total += v
                _put(ws, r, 3 + dd, v, "#,##0")
            _put(ws, r, 4 + ndays, total, "#,##0", font=F_SUM)
            r += 1
        # 브랜드 셀 세로 병합
        ws.merge_cells(start_row=start, start_column=2, end_row=r - 1, end_column=2)
        cell = ws.cell(start, 2, b)
        cell.font = F_SUM
        cell.alignment = center_v

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 18
    for c in range(4, 4 + ndays):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 9
    ws.column_dimensions[ws.cell(row=1, column=4 + ndays).column_letter].width = 13

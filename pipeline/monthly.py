# -*- coding: utf-8 -*-
"""통합(다월) 리포트 전용: 브랜드별 '월별 매체 총 누적 + 증감' 시트.

각 월의 [매체 총 누적] 표(Total 시트와 동일 형태)를 세로로 나열하고,
연속 월 사이 전월 대비 증감율 표를 덧붙여 월별 변화를 보여준다.
단일월이면 생성하지 않음(build.py에서 판단)."""
from datetime import date
import period
from total import (media_cumulative, CUM_KEYS, CUM_FMT, _put, _div,
                   F_TITLE, F_SEC, F_COL, F_SUM, FILL_SEC, FILL_COL, FILL_SUM,
                   CENTER, LEFT, BRAND_TITLE)


def period_months():
    """리포트 기간 내 (연,월) 리스트 (오름차순)."""
    out, d = [], period.start()
    while d <= period.end():
        out.append((d.year, d.month))
        d = date(d.year + (d.month == 12), (d.month % 12) + 1, 1)
    return out


def _hdr_row(ws, r):
    _put(ws, r, 2, "구분", font=F_COL, fill=FILL_COL, align=CENTER)
    _put(ws, r, 3, "매체별 성과", font=F_COL, fill=FILL_COL, align=CENTER)
    for i, h in enumerate(CUM_KEYS):
        _put(ws, r, 4 + i, h, font=F_COL, fill=FILL_COL, align=CENTER)


def _cum_table(ws, r, title, cum_rows, total):
    """[매체 총 누적] 표 1개 (Total 시트와 동일 형태). 다음 시작행 반환."""
    _put(ws, r, 2, title, font=F_SEC, fill=FILL_SEC); r += 1
    _hdr_row(ws, r); r += 1
    for gubun, label, media, pat, budget, m in cum_rows:
        _put(ws, r, 2, gubun, align=CENTER)
        _put(ws, r, 3, label, align=LEFT)
        for i, k in enumerate(CUM_KEYS):
            _put(ws, r, 4 + i, m[k], CUM_FMT[i])
        if m["집행예산"] == 0:                      # 그 달 집행 0이면 숨김
            ws.row_dimensions[r].hidden = True
        r += 1
    _put(ws, r, 2, "합계", font=F_SUM, fill=FILL_SUM, align=CENTER)
    _put(ws, r, 3, "", font=F_SUM, fill=FILL_SUM)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    for i, k in enumerate(CUM_KEYS):
        _put(ws, r, 4 + i, total[k], CUM_FMT[i], font=F_SUM, fill=FILL_SUM)
    return r + 2


def _change_table(ws, r, title, prev, cur, ptot, ctot):
    """전월 대비 증감율 표: 각 지표 (cur-prev)/prev. prev=0·cur>0=신규, 둘다 0=숨김."""
    _put(ws, r, 2, title, font=F_SEC, fill=FILL_SEC); r += 1
    _hdr_row(ws, r); r += 1

    def putrow(gubun, label, pm, cm, sumrow=False):
        f = F_SUM if sumrow else None
        fl = FILL_SUM if sumrow else None
        _put(ws, r, 2, gubun, font=f, fill=fl, align=CENTER)
        _put(ws, r, 3, label, font=f, fill=fl, align=LEFT)
        for i, k in enumerate(CUM_KEYS):
            if pm[k] == 0 and cm[k] == 0:
                v = ""
            elif pm[k] == 0:
                v = "신규"
            else:
                v = _div(cm[k] - pm[k], pm[k])
            fmt = "0.0%" if isinstance(v, float) else None
            _put(ws, r, 4 + i, v, fmt, font=f, fill=fl,
                 align=None if isinstance(v, float) else CENTER)

    for (g, l, med, pat, bud, pm), (_, _, _, _, _, cm) in zip(prev, cur):
        putrow(g, l, pm, cm)
        if pm["집행예산"] == 0 and cm["집행예산"] == 0:
            ws.row_dimensions[r].hidden = True
        r += 1
    putrow("합계", "", ptot, ctot, sumrow=True)      # 먼저 쓰고
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)  # 그다음 병합
    return r + 2


def write_monthly_sheet(ws, brand, df_brand):
    """브랜드별 월별 [매체 총 누적] + 전월대비 증감율."""
    months = period_months()
    _put(ws, 2, 2, f"[ {BRAND_TITLE[brand]} ] 월별 매체 총 누적 · 증감", font=F_TITLE)
    r = 4
    cums = []
    for (y, m) in months:
        mdf = df_brand[(df_brand["날짜"].dt.year == y) & (df_brand["날짜"].dt.month == m)]
        cr, tot = media_cumulative(mdf, brand)
        cums.append((y, m, cr, tot))
        r = _cum_table(ws, r, f"■ {y}년 {m}월 [매체 총 누적]", cr, tot)
    # 전월 대비 증감율 (연속 월 쌍)
    for i in range(1, len(cums)):
        py, pm, pcr, ptot = cums[i - 1]
        cy, cm, ccr, ctot = cums[i]
        r = _change_table(ws, r, f"■ {pm}월 → {cm}월 증감율 (전월 대비)",
                          pcr, ccr, ptot, ctot)
    # 열폭 (Total 시트 [매체 총 누적]과 동일)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 16
    for c in range(4, 4 + len(CUM_KEYS)):
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 13


def write_monthly_sheets(book, df):
    """통합(다월)일 때만 브랜드별 월별누적 시트 생성. 단일월이면 아무것도 안 함."""
    if len(period_months()) < 2:
        return
    for brand in ["MI", "IT", "EBM"]:
        ws = book.create_sheet(f"{brand}_월별누적")
        write_monthly_sheet(ws, brand, df[df["브랜드"] == brand])

# -*- coding: utf-8 -*-
"""Phase 4: ●광고비집행현황 (다브랜드 일별 광고비 매트릭스).

일자 세로 × (시선전체/MI/EBM/IT) 블록. 각 블록: 전체·SA소계·DA소계 + 매체별.
값 = 광고비용 합계 (통합 기준).
"""
import warnings
warnings.simplefilter("ignore")
import calendar
from datetime import date
from total import (week_in_month, WD_KR, _put, F_TITLE, F_COL, F_SUM,
                   FILL_COL, FILL_SUM, SAT_COLOR, SUN_COLOR, CENTER)

SA_ITEMS = ["N검색", "D검색", "G검색"]
# (그룹헤더, 소계라벨, 브랜드(None=전체), DA항목들)
BLOCKS = [
    ("광고비용-시선 전체", "시선 전체", None,
     ["K디스", "G디스", "N디스", "크리테오", "RTB하우스", "I디스"]),
    ("광고비용-미샤", "MI 소계", "MI",
     ["K디스", "G디스", "N디스", "크리테오", "크루비", "RTB하우스", "I디스"]),
    ("광고비용-E.B.M", "E.B.M 소계", "EBM",
     ["K디스", "G디스", "N디스", "크리테오", "버즈빌", "I디스"]),
    ("광고비용-잇미샤", "IT 소계", "IT",
     ["K디스", "G디스", "N디스", "크리테오", "RTB하우스", "모비온", "I디스", "틱톡"]),
]


def _cat(d, cat):
    """매체 카테고리별 통합 서브셋. 미보유 네트워크(D검색/크루비 등)는 빈 셋."""
    m, camp = d["매체"], d["캠페인"]
    has_cpc = camp.str.contains("cpc", case=False, regex=False)
    if cat == "N검색":   return d[m == "Naver SA"]
    if cat == "G검색":   return d[(m == "Google") & has_cpc]
    if cat == "G디스":   return d[(m == "Google") & ~has_cpc]
    if cat == "K디스":   return d[m == "KKO"]
    if cat == "N디스":   return d[m == "Naver"]
    if cat == "크리테오": return d[m == "Criteo"]
    if cat == "RTB하우스": return d[m == "RTB"]
    if cat == "I디스":   return d[m == "Meta"]
    return d.iloc[0:0]   # D검색·크루비·버즈빌·모비온·틱톡 = 0


def write_exec_report(ws, uni, y, mth):
    ndays = calendar.monthrange(y, mth)[1]
    _put(ws, 2, 2, "■ 광고비용 집행현황", font=F_TITLE)
    for c, h in [(1, "일자"), (2, "요일"), (3, "주차")]:
        _put(ws, 6, c, h, font=F_COL, fill=FILL_COL, align=CENTER)

    blocks = []
    c = 4
    for gtitle, subtitle, brand, da_items in BLOCKS:
        dfb = uni if brand is None else uni[uni["브랜드"] == brand]
        item_daily = {it: _cat(dfb, it).groupby("날짜키")["광고비용"].sum().to_dict()
                      for it in SA_ITEMS + da_items}
        cols = ["전체", "소계"] + SA_ITEMS + ["소계"] + da_items
        _put(ws, 4, c, gtitle, font=F_COL, fill=FILL_COL, align=CENTER)
        _put(ws, 5, c, subtitle, font=F_COL, fill=FILL_COL, align=CENTER)
        for i, cn in enumerate(cols):
            _put(ws, 6, c + i, cn, font=F_COL, fill=FILL_COL, align=CENTER)
        blocks.append((c, da_items, item_daily))
        c += len(cols)

    def write_block_values(row, c0, da_items, item_daily, dk, bold=False):
        f = F_SUM if bold else None
        fl = FILL_SUM if bold else None
        get = (lambda it: sum(item_daily[it].values())) if dk is None \
            else (lambda it: item_daily[it].get(dk, 0))
        sa = sum(get(it) for it in SA_ITEMS)
        da = sum(get(it) for it in da_items)
        _put(ws, row, c0, sa + da, "#,##0", font=f, fill=fl)
        _put(ws, row, c0 + 1, sa, "#,##0", font=f, fill=fl)
        for i, it in enumerate(SA_ITEMS):
            _put(ws, row, c0 + 2 + i, get(it), "#,##0", font=f, fill=fl)
        _put(ws, row, c0 + 2 + len(SA_ITEMS), da, "#,##0", font=f, fill=fl)
        base = c0 + 3 + len(SA_ITEMS)
        for i, it in enumerate(da_items):
            _put(ws, row, base + i, get(it), "#,##0", font=f, fill=fl)

    # 누적 (row 7)
    _put(ws, 7, 1, "누적", font=F_SUM, fill=FILL_SUM)
    for c0, da_items, item_daily in blocks:
        write_block_values(7, c0, da_items, item_daily, None, bold=True)
    # 일자별 (row 8~)
    for day in range(1, ndays + 1):
        d = date(y, mth, day); dk = d.strftime("%Y%m%d"); row = 7 + day
        col = SAT_COLOR if d.weekday() == 5 else SUN_COLOR if d.weekday() == 6 else None
        _put(ws, row, 1, d, "yyyy-mm-dd", align=CENTER, color=col)
        _put(ws, row, 2, WD_KR[d.weekday()], align=CENTER, color=col)
        _put(ws, row, 3, week_in_month(d), align=CENTER)
        for c0, da_items, item_daily in blocks:
            write_block_values(row, c0, da_items, item_daily, dk)

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 5
    ws.column_dimensions["C"].width = 5

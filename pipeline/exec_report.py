# -*- coding: utf-8 -*-
"""Phase 4: ●광고비집행현황 (다브랜드 일별 광고비 매트릭스).

일자 세로 × (시선전체/MI/EBM/IT) 블록. 각 블록: 전체·SA소계·DA소계 + 매체별.
값 = 광고비용 합계 (통합 기준).
"""
import warnings
warnings.simplefilter("ignore")
import calendar
from datetime import date
from openpyxl.styles import PatternFill, Border, Side
from total import (week_in_month, WD_KR, _put, F_TITLE, F_COL, F_SUM,
                   FILL_COL, FILL_SUM, SAT_COLOR, SUN_COLOR, CENTER)

# 섹션 색: 전체(노랑) / SA(파랑) / DA(초록), 소계는 진하게
FILL_TOT = PatternFill("solid", fgColor="FFE699")
FILL_SA = PatternFill("solid", fgColor="DDEBF7")
FILL_SA_SUB = PatternFill("solid", fgColor="9DC3E6")
FILL_DA = PatternFill("solid", fgColor="E2EFDA")
FILL_DA_SUB = PatternFill("solid", fgColor="A9D08E")
_THICK = Side(style="medium", color="404040")

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


def _section_fills(da_items):
    """블록 컬럼별 색: [전체, SA소계, SA항목(무색)..., DA소계, DA항목...]"""
    return ([FILL_TOT, FILL_SA_SUB] + [None] * len(SA_ITEMS)
            + [FILL_DA_SUB] + [FILL_DA] * len(da_items))


def write_exec_report(ws, uni, y, mth):
    ndays = calendar.monthrange(y, mth)[1]
    last_row = 7 + ndays
    _put(ws, 2, 2, "■ 광고비용 집행현황", font=F_TITLE)
    for c, h in [(2, "일자"), (3, "요일"), (4, "주차")]:
        _put(ws, 6, c, h, font=F_COL, fill=FILL_COL, align=CENTER)

    blocks = []
    c = 5   # A열 비움 → B부터, 데이터블록은 E부터
    for gtitle, subtitle, brand, da_items in BLOCKS:
        dfb = uni if brand is None else uni[uni["브랜드"] == brand]
        item_daily = {it: _cat(dfb, it).groupby("날짜키")["광고비용"].sum().to_dict()
                      for it in SA_ITEMS + da_items}
        cols = ["전체", "SA소계"] + SA_ITEMS + ["DA소계"] + da_items
        fills = _section_fills(da_items)
        _put(ws, 4, c, gtitle, font=F_COL, fill=FILL_COL, align=CENTER)
        _put(ws, 5, c, subtitle, font=F_COL, fill=FILL_TOT, align=CENTER)
        # 5행 SA/DA 그룹 병합 헤더
        ws.merge_cells(start_row=5, start_column=c + 1, end_row=5, end_column=c + 4)
        _put(ws, 5, c + 1, "SA", font=F_COL, fill=FILL_SA_SUB, align=CENTER)
        ws.merge_cells(start_row=5, start_column=c + 5, end_row=5, end_column=c + len(cols) - 1)
        _put(ws, 5, c + 5, "DA", font=F_COL, fill=FILL_DA_SUB, align=CENTER)
        # 6행 컬럼헤더 (SA 상세: N/D/G검색은 색 없음)
        for i, cn in enumerate(cols):
            hf = None if i in (2, 3, 4) else fills[i]
            _put(ws, 6, c + i, cn, font=F_COL, fill=hf, align=CENTER)
        blocks.append((c, da_items, item_daily, fills))
        c += len(cols)
    end_col = c - 1

    def write_block_values(row, c0, da_items, item_daily, fills, dk, bold=False):
        f = F_SUM if bold else None
        get = (lambda it: sum(item_daily[it].values())) if dk is None \
            else (lambda it: item_daily[it].get(dk, 0))
        vals = [None]  # placeholder, filled below
        sa = sum(get(it) for it in SA_ITEMS)
        da = sum(get(it) for it in da_items)
        vals = [sa + da, sa] + [get(it) for it in SA_ITEMS] + [da] + [get(it) for it in da_items]
        for i, v in enumerate(vals):
            _put(ws, row, c0 + i, v, "#,##0", font=f, fill=fills[i])

    # 누적 (row 7)
    _put(ws, 7, 2, "누적", font=F_SUM, fill=FILL_SUM)
    for c0, da_items, item_daily, fills in blocks:
        write_block_values(7, c0, da_items, item_daily, fills, None, bold=True)
    # 일자별 (row 8~)
    for day in range(1, ndays + 1):
        d = date(y, mth, day); dk = d.strftime("%Y%m%d"); row = 7 + day
        col = SAT_COLOR if d.weekday() == 5 else SUN_COLOR if d.weekday() == 6 else None
        _put(ws, row, 2, d, "yyyy-mm-dd", align=CENTER, color=col)
        _put(ws, row, 3, WD_KR[d.weekday()], align=CENTER, color=col)
        _put(ws, row, 4, week_in_month(d), align=CENTER)
        for c0, da_items, item_daily, fills in blocks:
            write_block_values(row, c0, da_items, item_daily, fills, dk)

    # 블록 구분 테두리 (각 블록 시작열 왼쪽 + 마지막열 오른쪽)
    for row in range(4, last_row + 1):
        for c0, *_ in blocks:
            cell = ws.cell(row=row, column=c0)
            cell.border = Border(left=_THICK)
        rc = ws.cell(row=row, column=end_col)
        rc.border = Border(right=_THICK, left=rc.border.left)

    # 열 너비 (A 비움)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 5
    ws.column_dimensions["D"].width = 5
    for cc in range(5, end_col + 1):
        ws.column_dimensions[ws.cell(row=1, column=cc).column_letter].width = 11

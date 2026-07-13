# -*- coding: utf-8 -*-
"""전 시트 공통 디자인 마감: 글꼴 통일 + 얇은 테두리 + 정렬.

기존 채움색/폰트색(주말색 등)/집행현황 테두리는 보존하고 글꼴 패밀리만 통일.
"""
from openpyxl.styles import Font, Border, Side, Alignment

FONT = "맑은 고딕"
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
VCENTER = Alignment(vertical="center")

# 테두리 자동적용 제외 (자체 스타일 보유)
BORDER_SKIP = {"●광고비집행현황"}


def apply_global_style(book):
    for ws in book.worksheets:
        skip_border = ws.title in BORDER_SKIP
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                o = cell.font
                # 글꼴 패밀리만 통일, 볼드/크기/색 보존
                cell.font = Font(name=FONT, size=o.size or 10, bold=o.bold,
                                 italic=o.italic, color=o.color)
                # 정렬: 기존 수평정렬 보존 + 수직 가운데
                a = cell.alignment
                cell.alignment = Alignment(horizontal=a.horizontal,
                                           vertical="center", wrap_text=a.wrap_text)
                # 제목(14pt 이상)은 테두리 제외
                if not skip_border and (o.size or 10) < 14:
                    cell.border = BORDER
        if not skip_border:
            ws.sheet_view.showGridLines = False   # 테두리 있으니 격자선 숨김

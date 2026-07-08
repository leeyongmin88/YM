# -*- coding: utf-8 -*-
"""오케스트레이터: RAW → 통합 시트 생성 → 엑셀 저장.

실행:  python build.py
출력:  YM/output/통합_리포트.xlsx  (통합 시트)
"""
import warnings
warnings.simplefilter("ignore")
import pandas as pd
from config import OUT_DIR
from ingest import combine_ads
from ga import join_ga

# 통합 시트 컬럼 순서 (완성본과 동일, 17열)
UNIFIED_ORDER = [
    "날짜", "날짜키", "매체", "브랜드", "캠페인", "광고그룹", "광고(소재)",
    "광고비용", "노출수", "클릭수", "GA구매", "GA구매수익", "GA세션",
    "매핑상태", "매칭키", "회원가입수", "회원가입세션",
]


def build_unified():
    """통합 DataFrame (17열, 정렬 완료) 반환."""
    df = join_ga(combine_ads())
    df = df[UNIFIED_ORDER].copy()
    df = df.sort_values(["매체", "브랜드", "캠페인", "광고그룹", "광고(소재)", "날짜"]).reset_index(drop=True)
    return df


def save_excel(df, path, y=2026, mth=7):
    from total import write_total_sheet
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="yyyy-mm-dd") as xw:
        df.to_excel(xw, sheet_name="통합", index=False)
        ws = xw.sheets["통합"]
        for i, col in enumerate(df.columns, start=1):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = max(len(str(col)), 12) + 2
        # 날짜열(A) 형식 yyyy-mm-dd (시간 제거)
        for row in range(2, len(df) + 2):
            ws.cell(row=row, column=1).number_format = "yyyy-mm-dd"
        # Total 대시보드 3개 브랜드
        for brand in ["MI", "IT", "EBM"]:
            wsb = xw.book.create_sheet(f"{brand}_Total")
            write_total_sheet(wsb, brand, df[df["브랜드"] == brand], y, mth)
        # 매체별 상세 리포트
        from media import add_media_sheets
        add_media_sheets(xw.book, df, y, mth)
        # ●광고비집행현황
        from exec_report import write_exec_report
        write_exec_report(xw.book.create_sheet("●광고비집행현황"), df, y, mth)
        # 플랫표
        from flat import write_flat
        write_flat(xw.book.create_sheet("통합_캠페인일자별"), df, y, mth)
        # 브랜드 종합 + 리포트 추가 요청
        from summary import write_brand_summary, write_report_request
        write_brand_summary(xw.book.create_sheet("브랜드 종합"), df, y, mth)
        write_report_request(xw.book.create_sheet("리포트 추가 요청"), df, y, mth)
        _reorder_by_brand(xw.book)
    return path


def _reorder_by_brand(book):
    """시트를 브랜드 순서로 정렬 (참고파일 방식): 통합 → MI/EBM/IT 각 브랜드 블록."""
    brand_order = ["MI", "EBM", "IT"]
    suffix_order = ["Total", "N검색", "구글SA", "피맥스_리포트", "K디스", "크리테오",
                    "RTB", "메타_성과형", "메타_브랜딩형", "N디스"]
    desired = (["통합", "●광고비집행현황", "리포트 추가 요청", "통합_캠페인일자별", "브랜드 종합"]
               + [f"{b}_{s}" for b in brand_order for s in suffix_order])
    existing = {ws.title: ws for ws in book.worksheets}
    ordered = [existing[t] for t in desired if t in existing]
    ordered += [ws for ws in book.worksheets if ws not in ordered]
    book._sheets = ordered


def main():
    df = build_unified()
    out = OUT_DIR / "통합_리포트.xlsx"
    save_excel(df, out)
    # 요약
    print("통합 시트 생성 완료:", out)
    print("  총 행수:", len(df))
    g = df.groupby("매체")[["광고비용", "GA구매", "GA구매수익"]].sum()
    print(g.to_string())
    print("  광고비 합계: {:,.0f}".format(df["광고비용"].sum()))
    print("  GA매출 합계: {:,.0f}".format(df["GA구매수익"].sum()))
    print("  매핑상태:", dict(df["매핑상태"].value_counts()))


if __name__ == "__main__":
    main()

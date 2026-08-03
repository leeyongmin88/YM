# -*- coding: utf-8 -*-
"""오케스트레이터: RAW → 통합 시트 생성 → 엑셀 저장.

실행:  python build.py
출력:  YM/output/통합_리포트.xlsx  (통합 시트)
"""
import warnings
warnings.simplefilter("ignore")
import os
import calendar
from datetime import datetime, date
import pandas as pd
import config
from config import OUT_DIR
from ingest import combine_ads
from ga import join_ga
import period

# 통합 시트 컬럼 순서 (완성본과 동일, 17열)
UNIFIED_ORDER = [
    "날짜", "날짜키", "매체", "브랜드", "캠페인", "광고그룹", "광고(소재)",
    "광고비용", "노출수", "클릭수", "GA구매", "GA구매수익", "GA세션",
    "매핑상태", "매칭키", "회원가입수", "회원가입세션",
]


def month_folders():
    """YM_RAW이 세미콜론 다중 폴더면 리스트 반환(예: 6+7월 통합), 아니면 None."""
    raw = os.environ.get("YM_RAW", "")
    if ";" in raw:
        return [f.strip() for f in raw.split(";") if f.strip()]
    return None


def _read_folder(fol):
    """한 월 폴더(fol)를 RAW+GA+정액까지 읽어 통합 DataFrame 반환.
    모듈 전역(RAW_DIR·JEONGAEK·GA_DIR)을 해당 폴더로 재바인딩."""
    import ingest, ga
    p = config.YM_ROOT / fol
    config.RAW_DIR = p
    config.BUDGET_FILE = config._budget_file(p)
    ingest.RAW_DIR = p
    ingest.JEONGAEK = config._build_jeongaek()     # 그 달 정액 예산
    ga.GA_DIR = p / "GA"
    return join_ga(combine_ads())


def build_unified():
    """통합 DataFrame (17열, 정렬 완료) 반환. 다중 폴더면 월별로 읽어 concat."""
    folders = month_folders()
    if folders:
        df = pd.concat([_read_folder(f) for f in folders], ignore_index=True)
    else:
        df = join_ga(combine_ads())
    df = df[UNIFIED_ORDER].copy()
    df = df.sort_values(["매체", "브랜드", "캠페인", "광고그룹", "광고(소재)", "날짜"]).reset_index(drop=True)
    return df


def detect_period(df):
    """리포트 기간 = 데이터 최소월 1일 ~ 최대월 말일.
    단일월이면 그 달 전체(회귀 안전), 다월이면 연속 span(예: 6/1~7/31)."""
    dmin, dmax = df["날짜"].min().date(), df["날짜"].max().date()
    start = date(dmin.year, dmin.month, 1)
    end = date(dmax.year, dmax.month, calendar.monthrange(dmax.year, dmax.month)[1])
    return start, end


def save_excel(df, path, y=2026, mth=7):
    import total
    from total import write_total_sheet
    period.set_period(*detect_period(df))         # 날짜 로직 기준 기간 설정
    folders = month_folders()
    if folders:
        total.set_budget_folders(folders)         # 다월 예산 합산
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="yyyy-mm-dd") as xw:
        df.to_excel(xw, sheet_name="통합", index=False)
        ws = xw.sheets["통합"]
        for i, col in enumerate(df.columns, start=1):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = max(len(str(col)), 12) + 2
        # 날짜열(A) 형식 yyyy-mm-dd (시간 제거)
        for row in range(2, len(df) + 2):
            ws.cell(row=row, column=1).number_format = "yyyy-mm-dd"
        # 매체별 상세 리포트 (Total보다 먼저 → 광고비 합계셀 등록소 채움)
        from media import add_media_sheets
        add_media_sheets(xw.book, df, y, mth)
        # Total 대시보드 3개 브랜드 (등록된 셀 참조식 사용)
        for brand in ["MI", "IT", "EBM"]:
            wsb = xw.book.create_sheet(f"{brand}_Total")
            write_total_sheet(wsb, brand, df[df["브랜드"] == brand], y, mth)
        # ●광고비집행현황
        from exec_report import write_exec_report
        write_exec_report(xw.book.create_sheet("●광고비집행현황"), df, y, mth)
        # 플랫표
        from flat import write_flat
        write_flat(xw.book.create_sheet("통합_캠페인일자별"), df, y, mth)
        # 통합_ENG 캠페인 (메타 게시물참여 '(eng)' 전용 — Raw/Meta_ENG, 통합 로직과 분리)
        from eng import write_eng_sheet
        write_eng_sheet(xw.book.create_sheet("통합_ENG 캠페인"), y, mth)
        # 브랜드 종합 + 리포트 추가 요청
        from summary import write_brand_summary, write_report_request
        write_brand_summary(xw.book.create_sheet("브랜드 종합"), df, y, mth)
        write_report_request(xw.book.create_sheet("리포트 추가 요청"), df, y, mth)
        # 미맵핑 GA 점검 (매칭 안 되는 GA 원본 + 사유)
        from mapping import write_mapping_sheets
        write_mapping_sheets(xw.book, y, mth)
        # 전체 디자인 마감 (글꼴·테두리 통일)
        from style import apply_global_style
        apply_global_style(xw.book)
        _reorder_by_brand(xw.book)
    return path


def _reorder_by_brand(book):
    """시트를 브랜드 순서로 정렬 (참고파일 방식): 통합 → MI/EBM/IT 각 브랜드 블록."""
    brand_order = ["MI", "EBM", "IT"]
    suffix_order = ["Total", "N검색", "구글SA", "피맥스_리포트", "K디스", "크리테오",
                    "RTB", "메타_성과형", "메타_브랜딩형", "N디스"]
    desired = (["통합", "리포트 추가 요청", "●광고비집행현황", "브랜드 종합", "통합_캠페인일자별",
                "통합_ENG 캠페인",
                "미맵핑_분류", "미맵핑_광고매출", "미맵핑_광고가입", "미맵핑_NaverSA"]
               + [f"{b}_{s}" for b in brand_order for s in suffix_order])
    existing = {ws.title: ws for ws in book.worksheets}
    ordered = [existing[t] for t in desired if t in existing]
    ordered += [ws for ws in book.worksheets if ws not in ordered]
    book._sheets = ordered


def output_path():
    """출력 경로: output/통합_리포트_{기간}_생성{YYMMDD}.xlsx.
    기간 = period.label() (단일월 '2026년7월', 다월 '2026년6~7월').
    period가 먼저 설정돼 있어야 함(main에서 set). 같은 조합 재생성 시 _ver.N 누적."""
    ym = period.label().replace(" ", "")          # '2026년7월' 또는 '2026년6~7월'
    stamp = datetime.now().strftime("%y%m%d")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    base = OUT_DIR / f"통합_리포트_{ym}_생성{stamp}.xlsx"
    if not base.exists():
        return base
    n = 1
    while (OUT_DIR / f"통합_리포트_{ym}_생성{stamp}_ver.{n}.xlsx").exists():
        n += 1
    return OUT_DIR / f"통합_리포트_{ym}_생성{stamp}_ver.{n}.xlsx"


def detect_month(df):
    """RAW 데이터에서 대상 연·월 자동 감지 (가장 많은 날짜의 연월)."""
    ym = df["날짜"].dt.to_period("M")
    top = ym.value_counts().idxmax()
    return top.year, top.month


def main():
    df = build_unified()
    y, mth = detect_month(df)
    period.set_period(*detect_period(df))         # 기간 먼저 설정(파일명·라벨용)
    out = output_path()
    save_excel(df, out, y, mth)
    # 요약
    print(f"대상 기간: {period.label()}  ({period.start()} ~ {period.end()})")
    print("통합 시트 생성 완료:", out)
    print("  총 행수:", len(df))
    g = df.groupby("매체")[["광고비용", "GA구매", "GA구매수익"]].sum()
    print(g.to_string())
    print("  광고비 합계: {:,.0f}".format(df["광고비용"].sum()))
    print("  GA매출 합계: {:,.0f}".format(df["GA구매수익"].sum()))
    print("  매핑상태:", dict(df["매핑상태"].value_counts()))


if __name__ == "__main__":
    main()

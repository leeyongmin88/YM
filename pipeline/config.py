# -*- coding: utf-8 -*-
"""데일리 리포트 자동화 - 설정값
wb-ss-da-dailyreport-v14 스킬을 로컬 Python 파이프라인으로 구현.
"""
from pathlib import Path

# --- 경로 ---
# config.py는 pipeline/ 안 → 상위 폴더가 프로젝트 루트. 어디에 두든 자동 인식(이식성).
YM_ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = YM_ROOT / "Raw"
OUT_DIR = YM_ROOT / "output"
SAMPLE_REPORT = YM_ROOT / "Sample_Report" / "7월_260706.xlsx"

# --- 월예산 파일 (매월 사용자가 이 엑셀만 수정하면 집행율에 반영) ---
BUDGET_FILE = RAW_DIR / "예산.xlsx"


def load_budgets():
    """Raw/예산.xlsx → {(구분, 매체라벨): {"MI":, "EBM":, "IT":}}. 없으면 {} (코드 기본값 사용).
    헤더(1행): 구분 | 매체 | MI | EBM | IT. 2행부터 데이터."""
    if not BUDGET_FILE.exists():
        return {}
    from openpyxl import load_workbook
    ws = load_workbook(BUDGET_FILE, data_only=True).active
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None or row[1] is None:
            continue
        key = (str(row[0]).strip(), str(row[1]).strip())
        vals = {}
        for i, b in enumerate(("MI", "EBM", "IT")):
            v = row[2 + i] if len(row) > 2 + i else None
            try:
                vals[b] = float(v) if v not in (None, "") else 0.0
            except (TypeError, ValueError):
                vals[b] = 0.0
        out[key] = vals
    return out

# --- 광고비 보정 계수 (스킬 공통원칙) ---
# Meta·Google·Criteo = /0.9*1.1 , RTB·KKO = *1.1 , Naver 계열 그대로
COST_COEF = {
    "Meta":     1.1 / 0.9,
    "Google":   1.1 / 0.9,
    "Criteo":   1.1 / 0.9,
    "RTB":      1.1,
    "KKO":      1.1,
    "Naver":    1.0,
    "Naver SA": 1.0,
}

# --- 통합 시트 컬럼 순서 (A~Q, 17열) ---
UNIFIED_COLS = [
    "날짜", "날짜키", "매체", "브랜드", "캠페인", "광고그룹", "광고(소재)",
    "광고비용", "노출수", "클릭수", "GA구매", "GA구매수익", "GA세션",
    "매핑상태", "매칭키", "회원가입수", "회원가입세션",
]

# 유효 브랜드 (캠페인명 토큰 검증용)
BRANDS = {"MI", "IT", "EBM"}

# --- Naver SA 정액(고정비) 6계열: 월예산 ---
# 일별 광고비 = 월예산 / 해당월 일수, 데이터 존재 기간에 적용.
# 캠페인명 : (브랜드, 매칭키, 기본 월예산). 실제값은 예산파일(SA 네이버 브랜드검색/엠버서더형) 우선.
_JEONGAEK_DEFAULT = {
    "NAV_MI_SA_pf_bsa_정액":         ("MI",  "정액_MI_bsa",  9_240_000),
    "NAV_MI_SA_pf_Ambassador_정액":  ("MI",  "정액_MI_amb",  2_530_000),
    "NAV_IT_SA_pf_bsa_정액":         ("IT",  "정액_IT_bsa",  17_710_000),
    "NAV_IT_SA_pf_Ambassador_정액":  ("IT",  "정액_IT_amb",  3_157_000),
    "NAV_EBM_SA_pf_bsa_정액":        ("EBM", "정액_EBM_bsa", 1_540_000),
    "NAV_EBM_SA_pf_Ambassador_정액": ("EBM", "정액_EBM_amb", 1_892_000),
}


def _build_jeongaek():
    """정액 월예산: 예산파일(SA 네이버 브랜드검색/엠버서더형) 우선, 없으면 기본값.
    → 예산파일만 수정하면 정액 집행액·집행율이 함께 반영됨(단일 소스)."""
    bud = load_budgets()
    out = {}
    for camp, (brand, key, default) in _JEONGAEK_DEFAULT.items():
        label = "네이버 브랜드검색" if "bsa" in camp else "네이버 엠버서더형"
        out[camp] = (brand, key, bud.get(("SA", label), {}).get(brand, default))
    return out


JEONGAEK = _build_jeongaek()

# --- GA 조인 룩업 (추가정보 시트 기반, 매월 갱신 가능) ---
# KKO catalog: GA 세션캠페인(카탈로그) → KK코드
KKO_CATALOG = {
    "ebm_catalog": "KK0003",
    "it_catalog":  "KK0001",
    "mi_catalog":  "KK0002",
}

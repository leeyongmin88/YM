# -*- coding: utf-8 -*-
"""데일리 리포트 자동화 - 설정값
wb-ss-da-dailyreport-v14 스킬을 로컬 Python 파이프라인으로 구현.
"""
from pathlib import Path

# --- 경로 ---
YM_ROOT = Path(r"C:\Users\admin\YM")
RAW_DIR = YM_ROOT / "Raw"
OUT_DIR = YM_ROOT / "output"
SAMPLE_REPORT = YM_ROOT / "Sample_Report" / "7월_260706.xlsx"

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

# --- Naver SA 정액(고정비) 6계열: 월예산 (매월 사용자가 갱신) ---
# 일별 광고비 = 월예산 / 해당월 일수, 데이터 존재 기간에 적용.
# 캠페인명 : (브랜드, 매칭키, 월예산)
JEONGAEK = {
    "NAV_MI_SA_pf_bsa_정액":         ("MI",  "정액_MI_bsa",  9_240_000),
    "NAV_MI_SA_pf_Ambassador_정액":  ("MI",  "정액_MI_amb",  2_530_000),
    "NAV_IT_SA_pf_bsa_정액":         ("IT",  "정액_IT_bsa",  17_710_000),
    "NAV_IT_SA_pf_Ambassador_정액":  ("IT",  "정액_IT_amb",  3_157_000),
    "NAV_EBM_SA_pf_bsa_정액":        ("EBM", "정액_EBM_bsa", 1_540_000),
    "NAV_EBM_SA_pf_Ambassador_정액": ("EBM", "정액_EBM_amb", 1_892_000),
}

# --- GA 조인 룩업 (추가정보 시트 기반, 매월 갱신 가능) ---
# KKO catalog: GA 세션캠페인(카탈로그) → KK코드
KKO_CATALOG = {
    "ebm_catalog": "KK0003",
    "it_catalog":  "KK0001",
    "mi_catalog":  "KK0002",
}

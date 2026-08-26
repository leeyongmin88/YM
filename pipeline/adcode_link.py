# -*- coding: utf-8 -*-
"""애드코드 사전 연결 (별도 실행 · 3_애드코드연결.bat).

통합 데이터의 각 행에서 애드코드(MT/KK/CT/NG/DB/TT/TS/GS…)를 추출해
`애드코드사전.xlsx`의 속성(Brand·상품명·구분·최적화·타겟팅·기획전명·소재 등)을
열로 붙인 '시각화/피벗용 플랫표'를 만든다.

통합은 실행 시 RAW에서 새로 생성(build_unified) → 항상 최신(방법 A).
출력: output/통합_애드코드연결_{기간}_생성YYMMDD.xlsx
  - [통합_애드코드연결] 통합 + 애드코드 속성 (한 행 = 날짜×소재)
  - [미매칭코드] 통합엔 있으나 사전에 없는 코드 목록(사전 업데이트 안내용)
"""
import warnings
warnings.simplefilter("ignore")
import os
import re
from datetime import datetime
import pandas as pd
import config
from build import build_unified
import period

DICT_PATH = config.YM_ROOT / "애드코드사전.xlsx"
CODE_RE = re.compile(r"[A-Za-z]{2}\d{3,}")     # MT0001·KK0001·CT0010·NG0154 등


def load_dict():
    """애드코드사전.xlsx(1행 그룹헤더·2행 컬럼명·3행~ 데이터) → ({코드:{속성}}, 속성열순서).
    통합과 겹치는 '매체' 컬럼은 '코드_매체'로 회피."""
    if not DICT_PATH.exists():
        raise FileNotFoundError(f"애드코드 사전 파일이 없습니다: {DICT_PATH}")
    df = pd.read_excel(DICT_PATH, header=1, dtype=str).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    key = df.columns[0]                                    # '애드코드'
    df[key] = df[key].astype(str).str.strip().str.upper()
    df = df[(df[key] != "") & (df[key] != "NAN")]
    attr_src = [c for c in df.columns if c != key]
    ren = {c: ("코드_매체" if c == "매체" else c) for c in attr_src}
    df = df.rename(columns=ren)
    attr_cols = [ren[c] for c in attr_src]
    dic = {r[key]: {c: r[c] for c in attr_cols} for _, r in df.iterrows()}
    return dic, attr_cols


def extract_code(soje, matchkey, dict_codes):
    """광고(소재)·매칭키에서 애드코드 추출. 사전에 있는 코드를 우선 선택."""
    srcs = (str(soje), str(matchkey))
    for src in srcs:                                       # 1순위: 사전에 존재하는 코드
        for t in CODE_RE.findall(src):
            if t.upper() in dict_codes:
                return t.upper()
    for src in srcs:                                       # 2순위: 코드형 토큰(미매칭 표시용)
        toks = CODE_RE.findall(src)
        if toks:
            return toks[0].upper()
    return ""                                              # 코드 없음(검색광고 등)


def build_linked():
    """최신 통합 + 애드코드 속성 결합 DataFrame + (속성열, 사전코드집합) 반환."""
    df = build_unified().copy()
    dic, attr_cols = load_dict()
    codes = set(dic)
    df["애드코드"] = [extract_code(s, k, codes)
                   for s, k in zip(df["광고(소재)"], df["매칭키"])]
    for c in attr_cols:
        df[c] = df["애드코드"].map(lambda x: dic.get(x, {}).get(c, ""))
    df["코드매칭"] = df["애드코드"].map(
        lambda x: "매칭" if x in codes else ("미매칭" if x else "코드없음"))
    return df, attr_cols, codes


def _out_path(stamp):
    ym = period.label().replace(" ", "")
    config.OUT_DIR.mkdir(parents=True, exist_ok=True)
    base = config.OUT_DIR / f"통합_애드코드연결_{ym}_생성{stamp}.xlsx"
    if not base.exists():
        return base
    n = 1
    while (config.OUT_DIR / f"통합_애드코드연결_{ym}_생성{stamp}_ver.{n}.xlsx").exists():
        n += 1
    return config.OUT_DIR / f"통합_애드코드연결_{ym}_생성{stamp}_ver.{n}.xlsx"


def main():
    df, attr_cols, codes = build_linked()
    stamp = datetime.now().strftime("%y%m%d")
    out = _out_path(stamp)
    with pd.ExcelWriter(out, engine="openpyxl", datetime_format="yyyy-mm-dd") as xw:
        df.to_excel(xw, sheet_name="통합_애드코드연결", index=False)
        # 미매칭코드(사전에 없는 코드) → 사전 업데이트 안내
        um = df[df["코드매칭"] == "미매칭"]
        if len(um):
            _pre = re.compile(r"([A-Za-z]+)")
            _MED = {"MT": "Meta", "KK": "KKO", "CT": "Criteo", "NG": "Naver",
                    "DB": "Dable", "TT": "TikTok", "TS": "Toss",
                    "GS": "Google", "GP": "Google", "GG": "Google", "GY": "Google"}
            # ① 코드별 상세
            detail = (um.groupby("애드코드")
                        .agg(매체=("매체", "first"), 브랜드=("브랜드", "first"),
                             캠페인=("캠페인", "first"), 소재예시=("광고(소재)", "first"),
                             행수=("애드코드", "size"), 광고비=("광고비용", "sum"),
                             매출=("GA구매수익", "sum"))
                        .reset_index())
            detail["접두어"] = detail["애드코드"].map(
                lambda c: (_pre.match(str(c)).group(1).upper() if _pre.match(str(c)) else ""))
            detail = detail.sort_values(["매체", "애드코드"])
            detail = detail[["접두어", "애드코드", "매체", "브랜드", "캠페인",
                             "소재예시", "행수", "광고비", "매출"]]
            # ② 매체(접두어)별 요약
            summ = (detail.groupby("접두어")
                          .agg(추정매체=("매체", "first"), 코드종수=("애드코드", "nunique"),
                               광고비=("광고비", "sum"), 매출=("매출", "sum"))
                          .reset_index().sort_values("광고비", ascending=False))
            summ.to_excel(xw, sheet_name="미매칭_요약", index=False)
            detail.to_excel(xw, sheet_name="미매칭_코드목록", index=False)

    # ── 요약 출력 ──
    tot = len(df)
    n_match = (df["코드매칭"] == "매칭").sum()
    n_unmatch = (df["코드매칭"] == "미매칭").sum()
    n_none = (df["코드매칭"] == "코드없음").sum()
    cost_match = df.loc[df["코드매칭"] == "매칭", "광고비용"].sum()
    cost_all_coded = df.loc[df["코드매칭"] != "코드없음", "광고비용"].sum()
    print(f"대상 기간: {period.label()}")
    print("애드코드 연결 완료:", out)
    print(f"  총 {tot}행 | 매칭 {n_match} · 미매칭 {n_unmatch} · 코드없음(검색 등) {n_none}")
    if cost_all_coded:
        print(f"  코드有 광고비 중 사전매칭률: {cost_match / cost_all_coded * 100:.1f}%")
    if n_unmatch:
        miss = sorted(df.loc[df["코드매칭"] == "미매칭", "애드코드"].unique())
        print(f"  ⚠ 사전에 없는 코드 {len(miss)}종 → [미매칭_요약]·[미매칭_코드목록] 시트 확인, 사전 업데이트 권장")
        print("     예:", ", ".join(miss[:10]))
    print(f"  붙인 속성열: {', '.join(attr_cols)}")


if __name__ == "__main__":
    main()

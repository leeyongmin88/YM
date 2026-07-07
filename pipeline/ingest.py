# -*- coding: utf-8 -*-
"""Phase 1a: 매체별 RAW 파일 읽기 → 정규화 → 광고비 보정 → 매칭키 산출.

각 리더는 표준 스키마 DataFrame 반환:
  매체, 브랜드, 캠페인, 광고그룹, 광고(소재), 날짜(Timestamp),
  광고비_raw, 노출수, 클릭수, 매칭키
광고비용(보정) = 광고비_raw * COST_COEF[매체] 는 combine_ads()에서 적용.
"""
import re
import calendar
import warnings
warnings.simplefilter("ignore")   # openpyxl 기본스타일 경고 등 숨김
import pandas as pd
from openpyxl import load_workbook
from config import RAW_DIR, COST_COEF, BRANDS, JEONGAEK

STD = ["매체", "브랜드", "캠페인", "광고그룹", "광고(소재)", "날짜",
       "광고비_raw", "노출수", "클릭수", "매칭키"]


# ---------- 유틸 ----------
def to_num(x):
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace("\xa0", "").replace('"', "").replace(",", "").replace(" ", "").strip()
    if s == "" or s.lower() == "nan":
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def to_date(x):
    if isinstance(x, pd.Timestamp):
        return x.normalize()
    s = str(x).strip().rstrip(".")
    s = s.replace(".", "-").replace("/", "-")
    ts = pd.to_datetime(s, errors="coerce")
    return ts.normalize() if pd.notna(ts) else pd.NaT


def brand_from(campaign, idx=1):
    toks = re.split(r"[_/ ]", str(campaign))
    if len(toks) > idx and toks[idx] in BRANDS:
        return toks[idx]
    for t in toks:
        if t in BRANDS:
            return t
    return ""


def norm_id(x):
    """광고/캠페인 ID → Excel 15자리 유효숫자 정규화 문자열. Meta 매칭용."""
    v = to_num(x)
    if v <= 0:
        return ""
    return "%.15g" % v


def _code(text, prefix):
    """text에서 첫 'prefix\\d+' 코드 추출 (KK/NG/CT). 없으면 ''"""
    m = re.search(prefix + r"\d+", str(text))
    return m.group(0) if m else ""


def _xlsx_rows(path, sheet=None):
    wb = load_workbook(path, read_only=False, data_only=True)  # Criteo read_only 버그 회피
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows


def _read_csv(path, encoding, sep=",", skiprows=0):
    return pd.read_csv(path, encoding=encoding, sep=sep, skiprows=skiprows,
                       dtype=str, keep_default_na=False)


# ---------- 매체별 리더 (매칭키 포함) ----------
def read_meta():
    out = []
    for f in sorted((RAW_DIR / "Meta").glob("*.xlsx")):
        for r in _xlsx_rows(f, "Raw Data Report")[1:]:
            if r[0] is None:
                continue
            camp = str(r[1] or ""); cre = str(r[3] or "")
            key = _code(cre, "MT") or camp                    # 매칭키 = 소재의 MT코드
            out.append(["Meta", brand_from(camp), camp, str(r[2] or ""), cre,
                        to_date(r[0]), to_num(r[5]), to_num(r[6]), to_num(r[7]), key])
    return pd.DataFrame(out, columns=STD)


def read_google():
    f = next((RAW_DIR / "Google").glob("*.csv"))
    df = _read_csv(f, "utf-16", sep="\t", skiprows=2)
    df.columns = [c.strip() for c in df.columns]
    out = []
    for _, r in df.iterrows():
        camp = str(r["캠페인"]).strip()
        if not camp:
            continue
        out.append(["Google", brand_from(camp), camp, camp, camp,
                    to_date(r["일"]), to_num(r["비용"]), to_num(r["노출수"]), to_num(r["클릭수"]),
                    camp])                                    # 매칭키 = 캠페인명
    return pd.DataFrame(out, columns=STD)


def read_kko():
    f = next((RAW_DIR / "KKO").glob("*.csv"))
    df = _read_csv(f, "utf-16", sep="\t")
    df.columns = [c.strip().strip('"') for c in df.columns]
    out = []
    for _, r in df.iterrows():
        camp = str(r["캠페인 이름"]).strip()
        if not camp:
            continue
        cre = str(r["소재 이름"]).strip()
        key = _code(cre, "KK") or camp
        out.append(["KKO", brand_from(camp), camp, str(r["광고그룹 이름"]).strip(), cre,
                    to_date(r["일"]), to_num(r["비용"]), to_num(r["노출수"]), to_num(r["클릭수"]),
                    key])
    return pd.DataFrame(out, columns=STD)


def read_criteo():
    f = next((RAW_DIR / "Criteo").glob("*.xlsx"))
    out = []
    for r in _xlsx_rows(f, "Download")[1:]:
        if r[0] is None:
            continue
        camp = str(r[1] or ""); cre = str(r[5] or "")
        key = _code(cre, "CT") or camp
        out.append(["Criteo", brand_from(camp), camp, str(r[3] or ""), cre,
                    to_date(r[0]), to_num(r[6]), to_num(r[7]), to_num(r[8]), key])
    return pd.DataFrame(out, columns=STD)


def read_rtb():
    f = next((RAW_DIR / "RTB").glob("*.xlsx"))
    out = []
    for r in _xlsx_rows(f)[1:]:
        if r[0] is None:
            continue
        camp = str(r[0])                       # 'IT_pf'
        brand = brand_from(camp, idx=0)
        out.append(["RTB", brand, camp, camp, camp,
                    to_date(r[1]), to_num(r[5]), to_num(r[2]), to_num(r[3]),
                    brand])                                   # 매칭키 = 브랜드
    return pd.DataFrame(out, columns=STD)


def read_naver_advoost():
    f = next((RAW_DIR / "NAV").glob("*Advoost*.csv"))
    df = _read_csv(f, "utf-8-sig")
    df.columns = [c.strip() for c in df.columns]
    out = []
    for _, r in df.iterrows():
        camp = str(r["캠페인 이름"]).strip()
        if not camp:
            continue
        out.append(["Naver", brand_from(camp), camp, camp, camp, to_date(r["기간"]),
                    to_num(r["총비용"]), to_num(r["노출수"]), to_num(r["클릭수"]),
                    camp])                                    # advoost 매칭키 = 캠페인명
    return pd.DataFrame(out, columns=STD)


def read_naver_smart():
    f = next((RAW_DIR / "NAV").glob("*Smart*.csv"))
    df = _read_csv(f, "utf-8-sig")
    df.columns = [c.strip() for c in df.columns]
    out = []
    for _, r in df.iterrows():
        camp = str(r["캠페인 이름"]).strip()
        if not camp:
            continue
        cre = str(r["광고 소재 이름"]).strip()
        key = _code(cre, "NG") or camp
        out.append(["Naver", brand_from(camp), camp, str(r["광고 그룹 이름"]).strip(), cre,
                    to_date(r["기간"]), to_num(r["총비용"]), to_num(r["노출수"]), to_num(r["클릭수"]),
                    key])
    return pd.DataFrame(out, columns=STD)


def read_nsa():
    f = next((RAW_DIR / "NSA").glob("*.csv"))
    df = _read_csv(f, "utf-8-sig", skiprows=1)   # 1행 제목 skip
    df.columns = [c.strip() for c in df.columns]
    out = []
    for _, r in df.iterrows():
        camp = str(r["캠페인"]).strip()
        if not camp:
            continue
        out.append(["Naver SA", brand_from(camp), camp, camp, camp, to_date(r["일별"]),
                    to_num(r["총비용"]), to_num(r["노출수"]), to_num(r["클릭수"]),
                    camp])                                    # NSA 매칭키 = 캠페인(SA조인 별도)
    return pd.DataFrame(out, columns=STD)


READERS = [read_meta, read_google, read_kko, read_criteo, read_rtb,
           read_naver_advoost, read_naver_smart, read_nsa]


def build_jeongaek(date_min, date_max):
    days_in_month = calendar.monthrange(date_min.year, date_min.month)[1]
    dates = pd.date_range(date_min.replace(day=1), date_max, freq="D")
    out = []
    for camp, (brand, key, budget) in JEONGAEK.items():
        daily = budget / days_in_month
        for d in dates:
            out.append(["Naver SA", brand, camp, "정액", "정액", d.normalize(),
                        daily, 0.0, 0.0, key])
    return pd.DataFrame(out, columns=STD)


def combine_ads():
    parts = [fn() for fn in READERS]
    df = pd.concat(parts, ignore_index=True)
    df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce")
    df = df[df["날짜"].notna()].reset_index(drop=True)
    jg = build_jeongaek(df["날짜"].min(), df["날짜"].max())
    df = pd.concat([df, jg], ignore_index=True)
    df["날짜키"] = df["날짜"].dt.strftime("%Y%m%d")
    # 소재 단위 집계 (같은 소재가 여러 광고세트로 쪼개진 경우 합산 → GA 중복부여 방지)
    dims = ["매체", "브랜드", "캠페인", "광고그룹", "광고(소재)", "날짜", "날짜키", "매칭키"]
    df = (df.groupby(dims, as_index=False)
            .agg({"광고비_raw": "sum", "노출수": "sum", "클릭수": "sum"}))
    df["광고비용"] = df.apply(lambda r: r["광고비_raw"] * COST_COEF[r["매체"]], axis=1)
    return df


if __name__ == "__main__":
    df = combine_ads()
    g = df.groupby("매체")[["광고비용", "노출수", "클릭수"]].sum()
    print(g.to_string())
    print("\nrows:", len(df), " 광고비 합계:", f"{df['광고비용'].sum():,.0f}")

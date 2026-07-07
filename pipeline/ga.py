# -*- coding: utf-8 -*-
"""Phase 1b: GA 데이터 읽기 → 매칭키 산출 → (매칭키,날짜) 집계 → 통합 조인.

GA 원본 4파일:
  광고매출.csv        → 구매/구매수익/세션 (DA·검색 광고)
  광고가입.csv        → 회원가입(이벤트수)/가입세션
  네이버SA매출.csv    → Naver SA 구매/수익/세션
  네이버SA가입.csv    → Naver SA 회원가입/세션
모두 프리앰블 6줄 + 총합계행(헤더 다음 1줄) 존재 → skiprows=[0,1,2,3,4,5,7].
"""
import re
import pandas as pd
from config import RAW_DIR, KKO_CATALOG, BRANDS
from ingest import to_num, norm_id, _code

GA_DIR = RAW_DIR / "GA"
_SKIP = [0, 1, 2, 3, 4, 5, 7]   # 프리앰블 + 총합계행


def _read_ga(name):
    df = pd.read_csv(GA_DIR / name, encoding="utf-8-sig", skiprows=_SKIP,
                     dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]
    return df


def platform_of(source_medium):
    s = str(source_medium).lower()
    if "criteo" in s:    return "Criteo"
    if "rtbhouse" in s:  return "RTB"
    if "kakao" in s:     return "KKO"
    if s.startswith("naver"):  return "Naver"
    if s.startswith("google"): return "Google"
    if any(x in s for x in ["instagram", "igshopping", "fb /", "ig /"]): return "Meta"
    return ""


def ga_key(plat, camp, content, camp_id):
    """GA행 → 통합 매칭키 (매체별 규칙). 미매칭이면 ''"""
    camp = str(camp).strip()
    if plat == "Meta":
        return _code(content, "MT")                # GA 콘텐츠의 MT코드
    if plat == "Google":
        return camp if camp.upper().startswith("GGL") else ""
    if plat == "Criteo":
        return _code(camp_id, "CT")
    if plat == "RTB":
        return camp.split("_")[0].upper()          # it_rtb_re → IT
    if plat == "KKO":
        kk = _code(content, "KK")
        return kk or KKO_CATALOG.get(camp, "")
    if plat == "Naver":
        ng = _code(content, "NG")
        if ng:
            return ng
        if "advoost" in camp.lower() or "advoost" in str(content).lower():
            return "NAV_%s_DA_pf_advoost" % camp.split("_")[0].upper()
        return ""
    return ""


SA_DEVICE = {"mobile": "mo", "tablet": "mo", "desktop": "pc"}


def sa_campaign(sess_camp, device):
    """Naver SA GA행(세션캠페인+기기) → NSA 캠페인명 매칭키.
    예: it_brandsearch + mobile → NAV_IT_SA_pf_bsa_mo_new"""
    sc = str(sess_camp).strip().lower()
    brand = sc.split("_")[0].upper()
    if brand not in BRANDS:
        return ""                                   # ss_* 등 → 대상외
    dev = SA_DEVICE.get(str(device).strip().lower(), "mo")
    if "ambassador" in sc:
        return "NAV_%s_SA_pf_Ambassador" % brand
    if "brandsearch" in sc:
        return "NAV_%s_SA_pf_bsa_%s_new" % (brand, dev)
    if "paidsearch" in sc:
        return "NAV_%s_SA_pf_cpc_%s" % (brand, dev)
    if "shopping" in sc:
        return "NAV_%s_SA_pf_shopping_%s" % (brand, dev)
    return ""


def build_ga_sales():
    """(매칭키, 날짜키) → [구매, 구매수익, 세션]. DA/검색 + Naver SA 통합."""
    agg = {}
    # (1) DA·검색 광고매출
    df = _read_ga("광고매출.csv")
    df = df[df["세션 소스/매체"].str.strip() != ""]
    for _, r in df.iterrows():
        plat = platform_of(r["세션 소스/매체"])
        key = ga_key(plat, r["세션 캠페인"], r["세션 수동 광고 콘텐츠"], r["세션 캠페인 ID"])
        if not key:
            continue
        dk = str(r["날짜"]).strip()
        a = agg.setdefault((key, dk), [0.0, 0.0, 0.0])
        a[0] += to_num(r["구매"]); a[1] += to_num(r["구매 수익"]); a[2] += to_num(r["세션수"])
    # (2) Naver SA 매출
    sa = _read_ga("네이버SA매출.csv")
    sa = sa[sa["세션 소스/매체"].str.strip() != ""]
    dev_col = sa.columns[3]                          # 기기 카테고리
    for _, r in sa.iterrows():
        key = sa_campaign(r["세션 캠페인"], r[dev_col])
        if not key:
            continue
        dk = str(r["날짜"]).strip()
        a = agg.setdefault((key, dk), [0.0, 0.0, 0.0])
        a[0] += to_num(r["구매"]); a[1] += to_num(r["구매 수익"]); a[2] += to_num(r["세션수"])
    return agg


def build_ga_signup():
    """(매칭키, 날짜키) → [회원가입수, 회원가입세션]. DA/검색 + Naver SA."""
    agg = {}
    # (1) DA·검색 회원가입: 회원가입수=이벤트수, 세션=세션수
    df = _read_ga("광고가입.csv")
    df = df[df["세션 소스/매체"].str.strip() != ""]
    for _, r in df.iterrows():
        plat = platform_of(r["세션 소스/매체"])
        key = ga_key(plat, r["세션 캠페인"], r["세션 수동 광고 콘텐츠"], r["세션 캠페인 ID"])
        if not key:
            continue
        dk = str(r["날짜"]).strip()
        a = agg.setdefault((key, dk), [0.0, 0.0])
        a[0] += to_num(r["이벤트 수"]); a[1] += to_num(r["세션수"])
    # (2) Naver SA 회원가입 (세션수만 존재 → 가입수=세션=세션수)
    sa = _read_ga("네이버SA가입.csv")
    sa = sa[sa["세션 소스/매체"].str.strip() != ""]
    dev_col = sa.columns[3]
    for _, r in sa.iterrows():
        key = sa_campaign(r["세션 캠페인"], r[dev_col])
        if not key:
            continue
        dk = str(r["날짜"]).strip()
        n = to_num(r["세션수"])
        a = agg.setdefault((key, dk), [0.0, 0.0])
        a[0] += n; a[1] += n
    return agg


def join_ga(ad_df):
    """통합 광고행에 GA(구매/수익/세션 + 회원가입) 조인 + 매핑상태."""
    sales = build_ga_sales()
    signup = build_ga_signup()
    K, L, M, P, Q, status = [], [], [], [], [], []
    seen_s, seen_g = set(), set()          # (매칭키,날짜)당 GA 1회만 부여
    for _, r in ad_df.iterrows():
        key, dk, media = r["매칭키"], r["날짜키"], r["매체"]
        s = sales.get((key, dk)) if (key, dk) not in seen_s else None
        g = signup.get((key, dk)) if (key, dk) not in seen_g else None
        if s:
            seen_s.add((key, dk))
        if g:
            seen_g.add((key, dk))
        K.append(s[0] if s else 0.0)
        L.append(s[1] if s else 0.0)
        M.append(s[2] if s else 0.0)
        P.append(g[0] if g else 0.0)
        Q.append(g[1] if g else 0.0)
        matched = sales.get((key, dk)) is not None
        if media == "Naver SA" and r["광고그룹"] == "정액":
            status.append("정액(비용반영)")
        elif matched:
            status.append("GA 1:1 매칭")
        else:
            status.append("광고만(GA미발생)")
    out = ad_df.copy()
    out["GA구매"], out["GA구매수익"], out["GA세션"] = K, L, M
    out["회원가입수"], out["회원가입세션"] = P, Q
    out["매핑상태"] = status
    return out


if __name__ == "__main__":
    from ingest import combine_ads
    ad = combine_ads()
    uni = join_ga(ad)
    g = uni.groupby("매체")[["GA구매", "GA구매수익", "GA세션", "회원가입수"]].sum()
    print(g.to_string())
    print("\n매핑상태:", dict(uni["매핑상태"].value_counts()))

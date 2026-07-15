# -*- coding: utf-8 -*-
"""미맵핑 GA 점검: 매칭키가 안 잡히는 GA 원본 행 + 사유를 시트로 정리.
(스킬 wb-ss-da-dailyreport-v14 ② 미맵핑 단계 참고)

- 미맵핑_분류: 광고매출 기준 매핑됨/키미식별/매체대응없음 요약.
- 미맵핑_광고매출 / 미맵핑_광고가입 / 미맵핑_NaverSA: 미맵핑 원본 행 + 플랫폼 + 사유.
"""
import warnings
warnings.simplefilter("ignore")
from ga import _read_ga, platform_of, ga_key, sa_campaign
from ingest import to_num
from config import BRANDS
from total import _put, F_TITLE, F_COL, FILL_COL, F_SUM, FILL_SUM, CENTER, LEFT

# DA/검색: 플랫폼별 키 미식별 사유 (key=='' 이고 플랫폼 인식된 경우)
_DA_REASON = {
    "Meta":   "키 미식별 · 콘텐츠에 MT코드 없음",
    "Google": "키 미식별 · 캠페인이 GGL_ 로 시작 안 함",
    "Criteo": "키 미식별 · 캠페인ID에 CT코드 없음",
    "KKO":    "키 미식별 · KK코드 없음 & 카탈로그 매핑 아님",
    "Naver":  "키 미식별 · NG코드 없음 & advoost 아님",
    "RTB":    "키 미식별",
}


def _da_reason(plat):
    if not plat:
        return "매체 대응없음 · 소스/매체 미인식(오가닉·AI·기타 유입)"
    return _DA_REASON.get(plat, "키 미식별")


def _nsa_reason(sess_camp):
    brand = str(sess_camp).strip().split("_")[0].upper()
    if brand not in BRANDS:
        return "대상외 · 브랜드 토큰 아님(ss_ 등)"
    return "유형 미식별 · bsa/cpc/shopping/ambassador 아님"


def _collect_da(name, metrics):
    """광고매출/광고가입 미맵핑 행 수집 + 요약(매핑됨/키미식별/매체대응없음)."""
    df = _read_ga(name)
    df = df[df["세션 소스/매체"].str.strip() != ""]
    has_cid = "세션 캠페인 ID" in df.columns
    rows, summ = [], {"매핑됨": [0, 0.0, 0.0, 0.0], "키 미식별": [0, 0.0, 0.0, 0.0],
                      "매체 대응없음": [0, 0.0, 0.0, 0.0]}
    for _, r in df.iterrows():
        plat = platform_of(r["세션 소스/매체"])
        cid = r["세션 캠페인 ID"] if has_cid else ""
        key = ga_key(plat, r["세션 캠페인"], r["세션 수동 광고 콘텐츠"], cid)
        vals = [to_num(r.get(m, 0)) for m in metrics]
        sess = to_num(r.get("세션수", 0))
        grp = "매핑됨" if key else ("키 미식별" if plat else "매체 대응없음")
        s = summ[grp]
        s[0] += 1; s[1] += sess
        if len(metrics) >= 2:              # 매출류(구매/수익)만 요약에 누적
            s[2] += vals[0]; s[3] += vals[1] if len(vals) > 1 else 0
        if key:
            continue
        rows.append([r["세션 소스/매체"], r["세션 캠페인"], r["세션 수동 광고 콘텐츠"],
                     cid, r["날짜"], *vals, plat or "(미인식)", _da_reason(plat)])
    return rows, summ


def _collect_nsa(name, metrics):
    df = _read_ga(name)
    df = df[df["세션 소스/매체"].str.strip() != ""]
    dev_col = df.columns[3]
    rows = []
    for _, r in df.iterrows():
        key = sa_campaign(r["세션 캠페인"], r[dev_col])
        if key:
            continue
        vals = [to_num(r.get(m, 0)) for m in metrics]
        rows.append([r["세션 소스/매체"], r["세션 캠페인"], r[dev_col], r["날짜"],
                     *vals, _nsa_reason(r["세션 캠페인"])])
    return rows


def _table(ws, title, headers, rows, fmts, top=2):
    """A열 비움·B2 제목·B4 헤더·B5~ 데이터. 미집행 정렬(사유별)."""
    _put(ws, top, 2, title, font=F_TITLE)
    hr = top + 2
    for i, h in enumerate(headers):
        _put(ws, hr, 2 + i, h, font=F_COL, fill=FILL_COL, align=CENTER)
    for j, row in enumerate(rows):
        for i, v in enumerate(row):
            fmt = fmts[i] if i < len(fmts) else None
            _put(ws, hr + 1 + j, 2 + i, v, fmt,
                 align=LEFT if fmt is None else None)
    ws.column_dimensions["A"].width = 2
    return hr + len(rows)


def write_mapping_sheets(book, y, mth):
    # 광고매출 (구매/수익/세션)
    sales_rows, sales_sum = _collect_da("광고매출.csv", ["구매", "구매 수익", "세션수"])
    # 광고가입 (이벤트수=회원가입/세션)
    signup_rows, _ = _collect_da("광고가입.csv", ["이벤트 수", "세션수"])
    # 네이버SA매출 (구매/수익/세션)
    nsa_rows = _collect_nsa("네이버SA매출.csv", ["구매", "구매 수익", "세션수"])

    # ── 미맵핑_분류 (광고매출 기준 요약) ──
    ws = book.create_sheet("미맵핑_분류")
    _put(ws, 2, 2, "■ GA 미맵핑 분류 (광고매출 기준)", font=F_TITLE)
    for i, h in enumerate(["구분", "GA행수", "세션수", "구매", "구매수익"]):
        _put(ws, 4, 2 + i, h, font=F_COL, fill=FILL_COL, align=CENTER)
    order = ["매핑됨", "키 미식별", "매체 대응없음"]
    tot = [0, 0.0, 0.0, 0.0]
    for k, k2 in enumerate(order):
        s = sales_sum[k2]
        _put(ws, 5 + k, 2, k2, align=LEFT)
        for c, (v, f) in enumerate(zip(s, ["#,##0", "#,##0", "#,##0", "#,##0"])):
            _put(ws, 5 + k, 3 + c, v, f)
            tot[c] += v
    _put(ws, 8, 2, "합계", font=F_SUM, fill=FILL_SUM)
    for c, v in enumerate(tot):
        _put(ws, 8, 3 + c, v, "#,##0", font=F_SUM, fill=FILL_SUM)
    ws.column_dimensions["A"].width = 2

    da_hdr = ["세션 소스/매체", "세션 캠페인", "콘텐츠", "세션 캠페인ID", "날짜",
              "구매", "구매수익", "세션수", "플랫폼", "미맵핑 사유"]
    da_fmt = [None, None, None, None, None, "#,##0", "#,##0", "#,##0", None, None]
    _table(book.create_sheet("미맵핑_광고매출"),
           "■ 미맵핑 GA 원본 · 광고매출", da_hdr, sales_rows, da_fmt)

    su_hdr = ["세션 소스/매체", "세션 캠페인", "콘텐츠", "세션 캠페인ID", "날짜",
              "회원가입", "세션수", "플랫폼", "미맵핑 사유"]
    su_fmt = [None, None, None, None, None, "#,##0", "#,##0", None, None]
    _table(book.create_sheet("미맵핑_광고가입"),
           "■ 미맵핑 GA 원본 · 광고가입(회원가입)", su_hdr, signup_rows, su_fmt)

    nsa_hdr = ["세션 소스/매체", "세션 캠페인", "기기", "날짜",
               "구매", "구매수익", "세션수", "미맵핑 사유"]
    nsa_fmt = [None, None, None, None, "#,##0", "#,##0", "#,##0", None]
    _table(book.create_sheet("미맵핑_NaverSA"),
           "■ 미맵핑 GA 원본 · 네이버SA", nsa_hdr, nsa_rows, nsa_fmt)

    # 미맵핑 시트는 점검용 → 숨김 처리(파일엔 있으나 탭 미표시)
    for name in ("미맵핑_분류", "미맵핑_광고매출", "미맵핑_광고가입", "미맵핑_NaverSA"):
        book[name].sheet_state = "hidden"

# -*- coding: utf-8 -*-
"""리포트 대상 기간(period) 추상화 — 단일 월 또는 연속 다월(예: 6/1~7/31) 지원.

build.py가 시작 시 set_month() 또는 set_period()로 1회 설정하고,
날짜 로직(daily_frame·weekday_avg·주간·제목)이 이 모듈을 참조한다.
단일 월이면 시작=1일·끝=말일이라 기존 월간 동작과 100% 동일(회귀 안전)."""
import calendar
from datetime import date, timedelta

_START = date(2026, 7, 1)          # 기본값(안전) — build.py가 덮어씀
_END = date(2026, 7, 31)


def set_month(y, mth):
    """단일 월 기간."""
    global _START, _END
    _START = date(y, mth, 1)
    _END = date(y, mth, calendar.monthrange(y, mth)[1])


def set_period(start, end):
    """임의 연속 기간(start~end, 양끝 포함)."""
    global _START, _END
    _START, _END = start, end


def start():
    return _START


def end():
    return _END


def days():
    """기간 내 모든 날짜 리스트 (start~end)."""
    out, d = [], _START
    while d <= _END:
        out.append(d)
        d += timedelta(days=1)
    return out


def _monday(d):
    return d - timedelta(days=d.weekday())


def week_of(d):
    """기간 시작 기준 주차(1..N, 월요일 시작). 단일월이면 월내 주차와 동일."""
    return (_monday(d) - _monday(_START)).days // 7 + 1


def n_weeks():
    """기간 내 주차 수."""
    return week_of(_END)


def week_periods():
    """{주차: 'mm/dd~mm/dd'} (1..n_weeks). 주간현황 기간 라벨용."""
    out = {}
    for wk in range(1, n_weeks() + 1):
        ds = [d for d in days() if week_of(d) == wk]
        out[wk] = (f"{ds[0].month:02d}/{ds[0].day:02d}~{ds[-1].month:02d}/{ds[-1].day:02d}"
                   if ds else "")
    return out


def week_month(wk):
    """주차 대표 월(그 주 첫 날짜의 월) — '{월}월 {주}주' 라벨용."""
    ds = [d for d in days() if week_of(d) == wk]
    return ds[0].month if ds else _START.month


def is_single_month():
    return _START.year == _END.year and _START.month == _END.month


def label():
    """제목용 라벨: 단일월 '2026년 7월', 다월 '2026년 6~7월'."""
    if is_single_month():
        return f"{_START.year}년 {_START.month}월"
    if _START.year == _END.year:
        return f"{_START.year}년 {_START.month}~{_END.month}월"
    return f"{_START.year}.{_START.month:02d}~{_END.year}.{_END.month:02d}"

# YM 데일리 광고 리포트 자동화

매체별 RAW 데이터를 폴더에 넣고 실행하면 **통합 리포트(.xlsx, 32시트)** 를 자동 생성하는 Python 파이프라인.
claude.ai 스킬 `wb-ss-da-dailyreport-v14`(Office.js 기반)를 로컬 Python으로 이식한 것.

## 빠른 시작

1. **RAW 파일 넣기** — `Raw/` 아래 매체별 폴더에 로우 파일을 넣는다.
   `Criteo · GA · Google · KKO · Meta · NAV · NSA · RTB`
2. **실행** — `1_리포트_생성.bat` 더블클릭.
3. **결과** — `output/통합_리포트.xlsx` 자동 생성.

> ⚠️ 실행 전 `output/통합_리포트.xlsx`가 엑셀에 **열려있으면 안 됨**(잠김). 대상 연·월은 데이터에서 자동 감지됨.

## 요구 사항

- Python 3.12
- `pip install pandas openpyxl`

수동 실행:
```powershell
$env:PYTHONIOENCODING="utf-8"
& "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe" "pipeline\build.py"
```

## 파이프라인 구성 (`pipeline/`)

| 파일 | 역할 |
|------|------|
| `config.py` | 경로, 광고비 보정계수, 정액 월예산, GA 룩업 설정 |
| `ingest.py` | 8개 매체 RAW → 정규화 → 광고비 보정 → 정액 → 매칭키 |
| `ga.py` | GA 조인 (매출·구매·세션·회원가입, 매칭키×날짜 1:1) |
| `build.py` | 오케스트레이터: 통합 시트 생성 + 대상월 자동감지 + 엑셀 저장 |
| `total.py` | 브랜드 Total 대시보드(MI/EBM/IT) + 공용 지표·스타일 |
| `media.py` | 매체별 상세 시트 (구글SA·피맥스·K디스·크리테오·RTB·메타·N검색·N디스) |
| `exec_report.py` | ●광고비집행현황 (일자세로 × 4브랜드 블록) |
| `flat.py` | 통합_캠페인일자별 (DA 실집행 플랫표) |
| `summary.py` | 브랜드 종합 + 리포트 추가 요청 |
| `style.py` | 전 시트 디자인 마감 (글꼴·테두리·병합·열너비) |

## 출력물

`output/통합_리포트.xlsx` — 32시트
(통합 · ●광고비집행현황 · 리포트 추가 요청 · 브랜드 종합 · 통합_캠페인일자별 + MI/EBM/IT 브랜드별 블록)

## 주의

- `Raw/`, `Sample_Report/`, `output/`는 실제 광고·매출 데이터라 `.gitignore`로 제외됨(코드만 공유, 데이터는 로컬 보관).
- 정액 월예산은 현재 `config.py`에 하드코딩 → 매월 갱신 필요(예산 파일 연동 예정).

# YM 데일리 광고 리포트 자동화

매체별 RAW 데이터를 넣으면 통합 리포트(.xlsx)를 자동 생성하는 파이프라인.
claude.ai 스킬 `wb-ss-da-dailyreport-v14`(Office.js 기반)를 로컬 Python으로 이식.

## 사용법

1. `Raw/` 폴더에 매체별 로우 파일을 넣는다 (Meta, Google, KKO, Criteo, RTB, NAV, NSA, GA).
2. `1_통합_확인.bat` 더블클릭 (또는 아래 명령 실행).

```powershell
$env:PYTHONIOENCODING="utf-8"; & "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe" "pipeline\ingest.py"
```

## 구성

- `pipeline/config.py` — 경로, 광고비 보정계수, 정액 예산, GA 룩업 설정
- `pipeline/ingest.py` — 8개 매체 RAW 읽기 → 정규화 → 광고비 보정 → 정액 → 매칭키
- `pipeline/ga.py` — GA 조인 (매출·구매·세션·회원가입, 매핑상태)

## 진행 상황

- ✅ Phase 1a: 광고 데이터 통합 (광고비 총액 완성본과 일치)
- 🔶 Phase 1b: GA 조인 (DA 매체 대부분 검증, Naver SA·정밀화 진행중)
- ⏳ Phase 2~: Total 대시보드, 매체별 리포트, 집행현황, 플랫표

## 주의

`Raw/`, `Sample_Report/`, `output/`는 실제 광고·매출 데이터라 `.gitignore`로 제외됨 (로컬 보관).
필요 라이브러리: `pip install pandas openpyxl`

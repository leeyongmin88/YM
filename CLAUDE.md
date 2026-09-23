# WB 데일리 광고 리포트 자동화 — 프로젝트 지침

매체별 RAW 데이터 → **통합 리포트(.xlsx, 32시트)** 를 생성하는 Python 파이프라인.
claude.ai 스킬 `wb-ss-da-dailyreport-v14`(Office.js)를 로컬 Python으로 이식한 것.
원 스킬 명세는 `.claude/skills/wb-ss-da-dailyreport-v14/SKILL.md` 참고 (리포트 구조·지표 정의의 원본).

## 실행 환경

- Python 3.12 (`%LOCALAPPDATA%\Programs\Python\Python312\python.exe`) — PATH 미등록이라 **전체 경로로 실행**
- pandas 3.0.3 / openpyxl 3.1.5 (`requirements.txt`)
- 한글 출력 깨짐 방지: 실행 전 `$env:PYTHONIOENCODING="utf-8"`

```powershell
$env:PYTHONIOENCODING="utf-8"
$env:YM_RAW="Raw_3_2026_08"     # 대상 월 폴더 (미설정이면 Raw/)
& "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe" "pipeline\build.py"
```

일반 사용자는 `1_리포트_생성.bat` 더블클릭 (월 선택 메뉴 → `YM_RAW` 자동 설정).

⚠️ **`output/` 의 대상 xlsx가 엑셀에 열려 있으면 PermissionError.** 검증만 필요하면 다른 이름(`_v2`)으로 저장해 확인하고, 사용자가 "닫았다"고 하면 본 파일 재생성.

## 폴더 구조

| 경로 | 내용 | Git |
|---|---|---|
| `pipeline/` | 파이프라인 코드 | 추적 |
| `키워드_Raw/` | 키워드 리포트 (`merge_gsa_ga.py`, `merge_nsa_ga.py`) | 추적 |
| `Raw_N_YYYY_MM/` | 월별 RAW + `예산.xlsx` | **ignore** |
| `기준정보/` | `애드코드사전_*.xlsx` | **ignore** |
| `output/` | 생성된 리포트 | **ignore** |

경로는 전부 `config.YM_ROOT = Path(__file__).resolve().parent.parent` 기준 상대경로. **절대경로 하드코딩 금지** (폴더를 어디에 두든 동작해야 함).

## 파이프라인 (`pipeline/`)

| 파일 | 역할 |
|---|---|
| `config.py` | 경로, 광고비 보정계수, 정액 예산, 예산파일 로더 |
| `ingest.py` | 매체 RAW → 정규화 → 광고비 보정 → 정액 → 매칭키 |
| `ga.py` | GA 조인 (매출·구매·세션·회원가입, 매칭키×날짜 1:1) |
| `build.py` | 오케스트레이터 · 대상월 자동감지 · 시트순서 정렬 |
| `period.py` | 대상 기간 추상화 (단일 월 / 연속 다월) |
| `total.py` | 브랜드 Total 대시보드 (MI/EBM/IT) + 공용 지표·스타일 상수 |
| `media.py` | 매체별 상세 시트 (구글SA·피맥스·K디스·크리테오·RTB·메타·N검색·N디스) |
| `exec_report.py` | ●광고비집행현황 |
| `flat.py` | 통합_캠페인일자별 (DA 실집행 플랫표) |
| `summary.py` | 브랜드 종합 + 리포트 추가 요청 |
| `mapping.py` | 미맵핑 GA 점검 시트 (전부 `sheet_state=hidden`) |
| `adcode_link.py` | 애드코드 사전 연결 (`3_애드코드연결.bat`) |
| `style.py` | 전 시트 디자인 마감 (글꼴·테두리·병합·열너비) |

## 핵심 규칙 (어기면 티가 안 나게 깨짐)

1. **색상 상수는 반드시 8자리 불투명 `FFxxxxxx`.** 6자리(`"BFBFBF"`)로 주면 openpyxl이 알파 `00`(투명)으로 저장해 **엑셀이 테두리를 투명 렌더** → "테두리가 안 보인다" 증상. 과거 이것 때문에 오래 헤맴.
2. **표 테두리색 = 회색 `FFBFBFBF`** (`style.THIN`, `total._TB_SIDE`). 합계/평균 음영 `F2F2F2`.
3. **출력은 계산된 값**(수식 아님). 시트 간 셀참조 금지 — 재구성 시 `#REF!`.
4. **레이아웃은 B2부터** (A열·1행 비움).
5. **월예산 단일 소스 = `Raw_*/예산.xlsx`.** 열: `구분|매체|통합매체|패턴|MI|EBM|IT`. 숫자만 고치면 정액 집행액·Total 집행율·브랜드종합 집행율에 전부 반영. **한 줄 추가 = 매체 추가** (단 이미 RAW가 들어오는 매체만; 새 플랫폼은 `ingest.py`에 리더 추가 필요).
6. **집행액 0인 매체는 자동 숨김** — Total은 행 숨김, ●광고비집행현황은 **열** 숨김. 집행 발생 시 재생성으로 자동 해제(동적).
7. **대상 연월은 자동감지** (`build.detect_month()`, 통합 날짜 최빈값). 월이 바뀌어도 코드 수정 불필요.
8. 사용자가 엑셀에서 직접 고친 파일을 주면 **diff로 의도를 추출**해 코드에 반영. 재생성하면 직접 수정분은 덮어써지므로, 확정된 디자인은 반드시 코드화.

## GA 매칭키 규칙

Meta=소재 MT코드 · Google=GGL 캠페인명 · Criteo=세션 캠페인ID의 CT코드 · KKO=콘텐츠 KK코드(없으면 catalog→`config.KKO_CATALOG`) · Naver=콘텐츠 NG코드(없으면 advoost→`NAV_{brand}_DA_pf_advoost`) · RTB=세션 캠페인 브랜드.
광고행은 소재 단위 집계 후 GA를 **(매칭키, 날짜)당 1회만** 부여 (중복 방지).

## 데이터 함정 (재발견하지 말 것)

- **Criteo xlsx**: openpyxl `read_only=True`면 1행만 읽힘 → `read_only=False`.
- **RAW 하단 '합계' 행**: 날짜가 NaT → 제거하지 않으면 수치 2배.
- **GA CSV**: 프리앰블 6줄 + 총합계행(9필드 > 헤더 8필드 → pandas 인덱스 오인으로 열이 밀림) → `skiprows=[0,1,2,3,4,5,7]`.
- **소스/매체 매핑**: `rtbhouse`→RTB, `instagram`/`igshopping`/`fb`/`ig`→Meta, `naver`→Naver, `naver/paidsearch`는 별도 SA 파일.
- **엠버서더는 모바일 전용** → N검색 PC 섹션에서 제외, MO만.
- **N검색 지표 순서 quirk**: 세션수가 회원가입 **앞** (참고파일 그대로 맞춘 것, 버그 아님).

## 작업 방식

- 데이터 파일(xlsx/csv)은 `.gitignore` 대상. **커밋은 코드만.**
- 코드 수정 후에는 `build.py` 재실행으로 32시트 생성 및 광고비 총액 일치를 확인.
- 임의로 범위를 넓히지 말 것. 구조를 바꾸는 변경은 먼저 분석·보고하고 승인받을 것.

@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
set PYTHONIOENCODING=utf-8

echo ============================================
echo   YM 데일리 리포트 생성 - 월 선택
echo ============================================
echo.
echo   [0] 기본 Raw 폴더
set /a i=0
for /d %%D in ("%~dp0Raw_*") do (
    set /a i+=1
    set "F!i!=%%~nxD"
    echo   [!i!] %%~nxD
)
echo.
echo   * 여러 달 통합 리포트는 1+2 처럼 +로 연결 (예: 6월+7월 통합)
set /p CHOICE="  번호 선택 (그냥 Enter = 0, 기본 Raw): "
if "%CHOICE%"=="" set CHOICE=0

set "YM_RAW="
set "PICK="
for %%N in (%CHOICE:+= %) do (
    if "%%N"=="0" (
        set "PICK=!PICK! Raw(기본)"
    ) else (
        if defined F%%N (
            if defined YM_RAW ( set "YM_RAW=!YM_RAW!;!F%%N!" ) else ( set "YM_RAW=!F%%N!" )
            set "PICK=!PICK! !F%%N!"
        ) else (
            echo.
            echo   [오류] %%N 번 폴더가 없습니다. 번호를 확인하고 다시 실행해주세요.
            echo.
            pause
            exit /b
        )
    )
)

echo.
echo   ^> 대상:!PICK!
echo.
rem --- Python 3.12 찾기: 기본 위치 -> py 런처 순 ---
set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if exist "%PY%" goto :pyfound
for /f "delims=" %%P in ('py -3.12 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%P"
if exist "%PY%" goto :pyfound
echo.
echo   [오류] Python 3.12 를 찾지 못했습니다.
echo          설치안내.md 의 "2. Python 3.12 설치" 를 확인해주세요.
echo.
pause
exit /b
:pyfound
"%PY%" "%~dp0pipeline\build.py"
echo.
echo ============================================
echo   완료. output 폴더에서 파일을 확인하세요.
echo   (파일명에 대상 기간이 표시됩니다)
echo   아무 키나 누르면 닫힙니다.
echo ============================================
pause >nul

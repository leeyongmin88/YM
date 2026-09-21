@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

echo ============================================
echo   키워드 리포트 생성 - 월 선택
echo ============================================
echo.
set /a i=0
for /d %%D in ("%~dp0Raw_*") do (
    set /a i+=1
    set "F!i!=%%~nxD"
    echo   [!i!] %%~nxD
)
if %i%==0 (
    echo   [오류] Raw_N_YYYY_MM 형식의 월 폴더가 없습니다.
    echo          예: Raw_3_2026_08\NSA\ , Raw_3_2026_08\GSA\
    pause
    exit /b
)
echo.
set /p CHOICE="  번호 선택: "

if not defined F%CHOICE% (
    echo.
    echo   [오류] %CHOICE% 번 폴더가 없습니다.
    pause
    exit /b
)
set "TARGET=!F%CHOICE%!"

echo.
echo   ^> 대상: !TARGET!
if not exist "%~dp0output" mkdir "%~dp0output"

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

echo.
echo ---- [1/2] 네이버 SA 키워드 ----
if exist "%~dp0!TARGET!\NSA" (
    "%PY%" "%~dp0merge_nsa_ga.py" --dir "%~dp0!TARGET!\NSA" --out "%~dp0output\NSA_키워드매칭_!TARGET!.xlsx"
) else (
    echo   [건너뜀] !TARGET!\NSA 폴더가 없습니다.
)

echo.
echo ---- [2/2] 구글 SA 키워드 ----
if exist "%~dp0!TARGET!\GSA" (
    "%PY%" "%~dp0merge_gsa_ga.py" --dir "%~dp0!TARGET!\GSA" --out "%~dp0output\GSA_키워드매칭_!TARGET!.xlsx"
) else (
    echo   [건너뜀] !TARGET!\GSA 폴더가 없습니다.
)

echo.
echo ============================================
echo   완료. output 폴더에서 결과를 확인하세요.
echo   아무 키나 누르면 닫힙니다.
echo ============================================
pause >nul

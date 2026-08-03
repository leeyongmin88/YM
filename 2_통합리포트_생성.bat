@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
set PYTHONIOENCODING=utf-8

echo ============================================
echo   YM 통합 리포트 생성 (여러 달 합치기)
echo ============================================
echo.
echo   합칠 월 폴더 목록:
set /a i=0
for /d %%D in ("%~dp0Raw_*") do (
    set /a i+=1
    set "F!i!=%%~nxD"
    echo   [!i!] %%~nxD
)
echo.
echo   * 합칠 월 번호를 공백 또는 +로 입력하세요
echo     예) 1 2      = 6월+7월 통합
echo         1+2+3    = 6월+7월+8월 통합
echo         2 3      = 7월+8월 통합
set /p CHOICE="  포함할 월 번호: "

set "YM_RAW="
set "PICK="
set /a cnt=0
for %%N in (%CHOICE:+= %) do (
    if defined F%%N (
        if defined YM_RAW ( set "YM_RAW=!YM_RAW!;!F%%N!" ) else ( set "YM_RAW=!F%%N!" )
        set "PICK=!PICK! !F%%N!"
        set /a cnt+=1
    ) else (
        echo.
        echo   [오류] %%N 번 폴더가 없습니다. 번호를 확인하고 다시 실행해주세요.
        echo.
        pause
        exit /b
    )
)
if %cnt%==0 (
    echo.
    echo   [오류] 최소 한 개 이상의 월을 선택해야 합니다.
    echo.
    pause
    exit /b
)

echo.
echo   ^> 통합 대상:!PICK!
echo   ^> 생성 중...
echo.
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0pipeline\build.py"
echo.
echo ============================================
echo   완료. output 폴더에서 파일을 확인하세요.
echo   파일명 예: 통합_리포트_2026년6~7월_생성YYMMDD.xlsx
echo   아무 키나 누르면 닫힙니다.
echo ============================================
pause >nul

@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
set PYTHONIOENCODING=utf-8

echo ============================================
echo   애드코드 사전 연결 - 월 선택
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
set /p CHOICE="  번호 선택 (그냥 Enter = 0, 기본 Raw): "
if "%CHOICE%"=="" set CHOICE=0

set "YM_RAW="
set "PICK="
if "%CHOICE%"=="0" (
    set "PICK=Raw(기본)"
) else (
    if defined F%CHOICE% (
        set "YM_RAW=!F%CHOICE%!"
        set "PICK=!F%CHOICE%!"
    ) else (
        echo.
        echo   [오류] %CHOICE% 번 폴더가 없습니다. 번호를 확인하고 다시 실행해주세요.
        echo.
        pause
        exit /b
    )
)

echo.
echo   ^> 대상: !PICK!
echo   ^> 애드코드사전.xlsx 기준으로 연결합니다.
echo.
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0pipeline\adcode_link.py"
echo.
echo ============================================
echo   완료. output 폴더에서 '통합_애드코드연결_...' 파일 확인.
echo   (사전에 없는 코드는 [미매칭코드] 시트에 표시됩니다)
echo   아무 키나 누르면 닫힙니다.
echo ============================================
pause >nul

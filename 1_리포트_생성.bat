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
set /p CHOICE="  번호 선택 (그냥 Enter = 0, 기본 Raw): "
if "%CHOICE%"=="" set CHOICE=0

if "%CHOICE%"=="0" (
    set "YM_RAW="
    set "PICK=Raw (기본)"
) else (
    set "YM_RAW=!F%CHOICE%!"
    set "PICK=!F%CHOICE%!"
    if not defined YM_RAW (
        echo.
        echo   [오류] 잘못된 번호입니다. 다시 실행해주세요.
        echo.
        pause
        exit /b
    )
    if not exist "%~dp0!YM_RAW!\" (
        echo.
        echo   [오류] !YM_RAW! 폴더를 찾을 수 없습니다.
        echo.
        pause
        exit /b
    )
)

echo.
echo   ^> !PICK! 폴더로 리포트를 생성합니다...
echo.
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0pipeline\build.py"
echo.
echo ============================================
echo   완료. output 폴더에서 파일을 확인하세요.
echo   (파일명에 데이터 월이 표시됩니다)
echo   아무 키나 누르면 닫힙니다.
echo ============================================
pause >nul

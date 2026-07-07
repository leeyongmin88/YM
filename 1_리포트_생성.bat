@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
echo ============================================
echo   YM Daily Report - Build Unified Sheet
echo ============================================
echo.
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0pipeline\build.py"
echo.
echo ============================================
echo   Done. Output file is in the output folder.
echo   Press any key to close.
echo ============================================
pause >nul
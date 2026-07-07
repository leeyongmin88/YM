@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
echo ============================================
echo   YM Daily Report - Unified Ad Data Check
echo ============================================
echo.
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0pipeline\ingest.py"
echo.
echo ============================================
echo   Done. Press any key to close.
echo ============================================
pause >nul
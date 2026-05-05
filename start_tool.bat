@echo off
title Amazon Ad Analyzer

echo ============================================
echo    Amazon Ad Analyzer - Starting...
echo ============================================
echo.

REM --- Find Python ---
set PY=
py --version >nul 2>&1 && set PY=py
if not defined PY (
    python -c "print(1)" >nul 2>&1 && set PY=python
)
if not defined PY (
    echo [ERROR] Python not found!
    echo Please install from https://www.python.org/downloads/
    echo Remember to check "Add Python to PATH"
    echo.
    pause
    exit /b
)

echo Python found:
%PY% --version
echo.

REM --- Install all dependencies at once ---
echo Installing dependencies (skip if already installed)...
%PY% -m pip install pandas openpyxl --quiet
echo.

REM --- Launch the tool ---
echo Launching tool...
echo.
%PY% "%~dp0amazon_ad_tool.py"
echo.

REM --- If we get here, something went wrong or user closed the window ---
if %errorlevel% neq 0 (
    echo.
    echo Something went wrong. The error message is shown above.
    echo.
)
pause

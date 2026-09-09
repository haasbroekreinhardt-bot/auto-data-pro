@echo off
REM ===================================================================
REM  AUTO-DATA PRO : EXCEL AUTOMATION DASHBOARD
REM  One-click Windows launcher.  Double-click this file to start.
REM ===================================================================
title Auto-Data Pro - Excel Automation Dashboard
cd /d "%~dp0"

echo.
echo  ==================================================================
echo    AUTO-DATA PRO : EXCEL AUTOMATION DASHBOARD
echo    Client: M^&M Retail   ^|   Rules Engine v2.4 ^(Active^)
echo  ==================================================================
echo.

REM ---- 1. Locate a Python interpreter ---------------------------------
set "PY_CMD="
where py >nul 2>&1 && set "PY_CMD=py -3"
if not defined PY_CMD (
    where python >nul 2>&1 && set "PY_CMD=python"
)
if not defined PY_CMD (
    echo  [ERROR] Python was not found on this machine.
    echo          Install Python 3.9 or newer from https://www.python.org/downloads/
    echo          and tick "Add python.exe to PATH" during setup.
    echo.
    pause
    exit /b 1
)
echo  [1/3] Python interpreter : %PY_CMD%

REM ---- 2. Ensure dependencies are installed ---------------------------
%PY_CMD% -c "import streamlit, pandas, openpyxl, xlsxwriter" >nul 2>&1
if errorlevel 1 (
    echo  [2/3] Installing dependencies ^(first run only, please wait^)...
    %PY_CMD% -m pip install --upgrade pip >nul 2>&1
    %PY_CMD% -m pip install -r "%~dp0requirements.txt"
    if errorlevel 1 (
        echo.
        echo  [ERROR] Dependency installation failed.
        echo          Try running this command manually:
        echo              %PY_CMD% -m pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
) else (
    echo  [2/3] Dependencies    : already installed
)

REM ---- Suppress Streamlit's first-run "enter your email" prompt --------
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit" >nul 2>&1
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    >"%USERPROFILE%\.streamlit\credentials.toml" echo [general]
    >>"%USERPROFILE%\.streamlit\credentials.toml" echo email = ""
)

REM ---- 3. Launch the dashboard -----------------------------------------
echo  [3/3] Starting Streamlit on http://localhost:8501
echo.
echo  The dashboard will open in your default browser.
echo  Keep this window open while you work - closing it stops the app.
echo.

%PY_CMD% -m streamlit run "%~dp0app.py" --server.port 8501 --server.headless false --browser.gatherUsageStats false --theme.base light

echo.
echo  Auto-Data Pro has stopped.
pause

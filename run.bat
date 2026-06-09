@echo off
REM ===========================================================================
REM  Agent QA Eval Suite - Windows launcher
REM  Double-click to run the eval suite locally. No make or Docker needed.
REM  First run sets things up (a minute); after that it starts in seconds.
REM ===========================================================================
cd /d "%~dp0"
title Agent QA Eval Suite

echo ============================================
echo   Agent QA Eval Suite
echo ============================================

REM 1) Python 3 --------------------------------------------------------------
set "PY="
where py  >nul 2>nul && set "PY=py"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY (
  echo [X] Python 3 is required but was not found.
  echo     Install from https://www.python.org/downloads/  ^(tick "Add python.exe to PATH"^)
  echo     then double-click this file again.
  pause & exit /b 1
)

REM 2) Azure CLI -------------------------------------------------------------
where az >nul 2>nul
if errorlevel 1 (
  echo [X] Azure CLI is required but was not found.
  echo     Install: https://learn.microsoft.com/cli/azure/install-azure-cli-windows
  echo     ^(or run: winget install Microsoft.AzureCLI^) then double-click this file again.
  pause & exit /b 1
)

REM 3) One-time setup: virtual env + dependencies ---------------------------
if not exist .venv (
  echo [*] First-time setup ^(installing - about a minute^)...
  %PY% -m venv .venv || ( echo [X] Could not create the environment. & pause & exit /b 1 )
  ".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
  ".venv\Scripts\python.exe" -m pip install --quiet -r requirements.txt || ( echo [X] Dependency install failed. & pause & exit /b 1 )
)

REM 4) Config ----------------------------------------------------------------
if not exist .env ( copy .env.example .env >nul & echo [*] Created .env from the template. )

REM 5) Azure sign-in (only if needed) ---------------------------------------
az account show >nul 2>nul
if errorlevel 1 (
  echo [*] Opening Azure sign-in in your browser...
  az login >nul || ( echo [X] Azure sign-in failed. & pause & exit /b 1 )
)

REM 6) Launch ----------------------------------------------------------------
set "PORT=8080"
for /f "tokens=2 delims==" %%a in ('findstr /b "APP_PORT=" .env') do set "PORT=%%a"
echo.
echo [*] Opening  http://localhost:%PORT%
echo     Keep this window open while you use the app. Close it to stop.
echo ============================================
start "" "http://localhost:%PORT%"
".venv\Scripts\python.exe" -m uvicorn app.main:app --host 127.0.0.1 --port %PORT%

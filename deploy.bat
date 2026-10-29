@echo off
rem ===========================================================================
rem  Office AI add-in: one-click deploy to C:\Program Files\OfficeAi
rem  REQUIREMENTS: Administrator (this script checks and tells you if not)
rem
rem  Why you must run this yourself:
rem    The agent session that prepared this runs unelevated, and Program Files
rem    only allows Administrators/SYSTEM/TrustedInstaller to write. That is a
rem    Windows security boundary, not a mistake in the command.
rem
rem  What it does:
rem    1) checks elevation        2) closes Word (the DLL is locked while loaded)
rem    3) backs up the current build
rem    4) copies new DLL + ui files + watchdog
rem    5) verifies MD5 of source vs Program Files, aborts on mismatch
rem
rem  NOTE: keep this file pure ASCII batch syntax (rem comments only).
rem        A PowerShell style block comment at the top makes cmd.exe fail with
rem        "The system cannot find the file specified".
rem ===========================================================================
setlocal EnableExtensions
title Office AI Add-in - Deploy

net session >nul 2>&1
if errorlevel 1 (
  echo.
  echo  [X] Not running as Administrator.
  echo.
  echo      Right-click this file  =^>  "Run as administrator"
  echo.
  echo      Or run this in an elevated PowerShell:
  echo        powershell -ExecutionPolicy Bypass -File "C:\Users\zhang\WorkBuddy\2026-10-15-03-49-08\office-ai\_dev\lead\elevated_deploy.ps1"
  echo.
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\zhang\WorkBuddy\2026-10-15-03-49-08\office-ai\_dev\lead\elevated_deploy.ps1"
set RC=%ERRORLEVEL%
echo.
if not "%RC%"=="0" (
  echo  [X] Deploy failed with exit code %RC%. Read the messages above.
) else (
  echo  [OK] Deploy finished. Reopen Word to load the new build.
)
echo.
pause
exit /b %RC%

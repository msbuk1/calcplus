@echo off
setlocal
title calcplus

REM Run from the folder that contains this script, wherever it sits.
cd /d "%~dp0"

where npm.cmd >nul 2>nul
if errorlevel 1 (
  echo.
  echo calcplus could not start: npm.cmd was not found.
  echo Install Node.js 24 LTS from https://nodejs.org/ and try again.
  echo.
  pause
  exit /b 1
)

echo Installing dependencies...
call npm.cmd install
if errorlevel 1 (
  echo.
  echo calcplus could not start: npm install failed.
  echo Read the errors above, fix them, and run this script again.
  echo.
  pause
  exit /b 1
)

echo.
echo Starting calcplus. Your browser should open automatically.
echo Keep this window open while you use calcplus.
echo Close this window to stop the server.
echo.
call npm.cmd run dev -- --open
if errorlevel 1 (
  echo.
  echo calcplus stopped: the dev server failed to start.
  echo Read the errors above, fix them, and run this script again.
  echo.
  pause
  exit /b 1
)

echo.
echo calcplus has stopped.
pause
exit /b 0

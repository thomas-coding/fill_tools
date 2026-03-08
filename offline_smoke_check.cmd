@echo off
setlocal

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found in PATH.
  echo Please install Python or run from an environment with python available.
  exit /b 2
)

if "%~1"=="" (
  echo Usage:
  echo   offline_smoke_check.cmd "path\to\source.xlsx"
  echo.
  echo Tip: You can drag and drop an Excel file onto this .cmd file.
  exit /b 1
)

python "%~dp0offline_smoke_check.py" "%~1" --samples 3
set "rc=%errorlevel%"

echo.
if not "%rc%"=="0" (
  echo Offline smoke check FAILED. Exit code: %rc%
  exit /b %rc%
)

echo Offline smoke check PASSED.
exit /b 0

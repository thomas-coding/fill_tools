@echo off
setlocal

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found in PATH.
  echo Please install Python or run from an environment with python available.
  exit /b 2
)

if "%~1"=="" (
  python "%~dp0release_preflight.py"
) else (
  python "%~dp0release_preflight.py" --sample "%~1"
)

exit /b %errorlevel%

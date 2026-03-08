@echo off
setlocal

set PY310=D:\software\eigent\resources\prebuilt\uv_python\cpython-3.10.19-windows-x86_64-none\python.exe
set BUILD_NAME=ShengDanTool

if not exist "%PY310%" (
  echo Python 3.10 runtime not found:
  echo %PY310%
  exit /b 1
)

pushd "%~dp0"

echo.
echo Running release preflight checks...
"%PY310%" "release_preflight.py"

if errorlevel 1 (
  echo.
  echo Preflight checks failed. Build aborted.
  popd
  exit /b 1
)

"%PY310%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "%BUILD_NAME%" ^
  --icon "app_icon.ico" ^
  --version-file "version_info.txt" ^
  --add-data "app_icon.ico;." ^
  --add-data "app_icon_preview.png;." ^
  --add-data "AutoHotkey-v2\AutoHotkey64.exe;AutoHotkey-v2" ^
  --add-data "AutoHotkey-v2\AutoHotkey32.exe;AutoHotkey-v2" ^
  app_main.py

if errorlevel 1 (
  echo.
  echo Build failed.
  popd
  exit /b 1
)

"%PY310%" -c "from pathlib import Path; src=Path(r'dist/ShengDanTool.exe'); dst=Path('dist')/'\u76db\u4e39\u7684\u5c0f\u5de5\u5177.exe'; dst.write_bytes(src.read_bytes())"

echo.
echo Build succeeded: dist\%BUILD_NAME%.exe
echo Added Chinese alias executable in dist folder.
popd
endlocal

@echo off
pushd "%~dp0"

python export_wechat_data.py
if errorlevel 1 (
  echo.
  echo Data export failed. Check wechat_form_test.xlsx and retry.
  pause
  popd
  exit /b 1
)

"AutoHotkey-v2\AutoHotkey64.exe" "wechat_form_helper.ahk"

python sync_progress_to_excel.py
if errorlevel 1 (
  echo.
  echo Progress sync failed.
)

popd

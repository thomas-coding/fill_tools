@echo off
pushd "%~dp0"

echo [1/3] 检查 Excel 进程...
tasklist /FI "IMAGENAME eq EXCEL.EXE" | find /I "EXCEL.EXE" >nul
if %errorlevel%==0 (
  echo 检测到 Excel 仍在运行，可能会导致文件只读。
  echo.
  set /p KILL_EXCEL=是否强制关闭所有 Excel 进程？(y/N): 
  if /I "%KILL_EXCEL%"=="Y" (
    taskkill /F /IM EXCEL.EXE >nul 2>&1
    timeout /t 1 >nul
  ) else (
    echo 已取消。请先手动关闭 Excel 后重试。
    pause
    popd
    exit /b 1
  )
)

echo [2/3] 清理锁文件...
if exist "~$wechat_form_test.xlsx" (
  del /f /q "~$wechat_form_test.xlsx"
)

echo [3/3] 完成。现在可编辑打开 wechat_form_test.xlsx。
pause
popd

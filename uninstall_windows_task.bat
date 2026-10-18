@echo off
chcp 65001 >nul
echo ========================================================
echo   移除 Windows 工作排程器任務
echo ========================================================
echo.

set TASK_NAME=CarbonFootprintDailyCrawler

echo 正在刪除排程任務 [%TASK_NAME%] ...
schtasks /delete /tn "%TASK_NAME%" /f

if %errorlevel% equ 0 (
    echo.
    echo [成功] 排程任務 [%TASK_NAME%] 已成功移除！
) else (
    echo.
    echo [提示] 任務可能不存在或請以「系統管理員身分執行」。
)
echo.
pause

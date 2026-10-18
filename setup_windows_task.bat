@echo off
chcp 65001 >nul
echo ========================================================
echo   註冊 Windows 工作排程器：每天凌晨 03:00 自動執行爬蟲
echo ========================================================
echo.

set TASK_NAME=CarbonFootprintDailyCrawler
set TASK_CMD="%~dp0run_task.bat"

echo 正在建立工作排程 [%TASK_NAME%] ...
schtasks /create /tn "%TASK_NAME%" /tr "%TASK_CMD%" /sc daily /st 03:00 /f

if %errorlevel% equ 0 (
    echo.
    echo ========================================================
    echo [成功] 已成功註冊 Windows 工作排程器！
    echo 任務名稱: %TASK_NAME%
    echo 執行時間: 每天凌晨 03:00
    echo 執行腳本: %TASK_CMD%
    echo 電腦開機狀態下，系統將在背景每天定時自動抓取並更新 data.xlsx！
    echo ========================================================
) else (
    echo.
    echo [提示] 若出現權限不足錯誤，請以「系統管理員身分執行」此批次檔。
)
echo.
pause

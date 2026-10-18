@echo off
chcp 65001 >nul
echo ========================================================
echo   綠色公益基金會 - 碳足跡排放係數資料庫 立即更新程式
echo ========================================================
echo.
cd /d "%~dp0"
echo 正在執行爬蟲並匯出至 data.xlsx ...
python crawler.py
echo.
if %errorlevel% equ 0 (
    echo [成功] 資料庫爬取完成，已成功儲存至 data.xlsx！
) else (
    echo [失敗] 爬取過程發生錯誤，請檢查網路連線或查看 crawler.log。
)
echo.
pause

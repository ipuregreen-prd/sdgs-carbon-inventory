@echo off
cd /d "%~dp0"
python crawler.py >> task_runner.log 2>&1

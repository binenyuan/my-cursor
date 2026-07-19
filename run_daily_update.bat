@echo off
REM Daily portfolio Excel auto-update (Windows)
REM Task Scheduler example: run at 15:30 on weekdays
cd /d "%~dp0"
echo ===== %date% %time% =====
python update_portfolio_excel.py
pause

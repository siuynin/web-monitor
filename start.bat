@echo off
chcp 65001 >nul
title Website Monitor
cd /d "%~dp0"
if not exist config.ini copy config.example.ini config.ini >nul
python -m pip install -q -r requirements.txt
:loop
python monitor.py
echo Monitor bi dung, khoi dong lai sau 10 giay...
timeout /t 10 >nul
goto loop

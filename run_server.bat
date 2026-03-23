@echo off
title SpareX Web Server
echo Starting SpareX Web Server...
echo.
echo Make sure your phone is connected to the same Wi-Fi.
echo.
cd /d "%~dp0"
python SpareX_Web.py
pause

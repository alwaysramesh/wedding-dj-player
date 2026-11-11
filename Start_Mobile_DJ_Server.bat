@echo off
title Ramesh Wedding DJ - Mobile & Laptop Server
color 0E
cls
echo ======================================================================
echo          RAMESH BROTHER WEDDING DJ PLAYER - MOBILE & WI-FI SERVER
echo                        10, 11 & 12 DECEMBER 2026
echo ======================================================================
echo.
echo  Detecting local Wi-Fi IP address...
for /f "tokens=4" %%a in ('route print 0.0.0.0 ^| findstr 0.0.0.0 ^| findstr /v "Default"') do (
    set LOCAL_IP=%%a
)
echo.
echo  ----------------------------------------------------------------------
echo   1. LAPTOP ACCESS:   http://localhost:8080
echo.
echo   2. MOBILE ACCESS:   http://%LOCAL_IP%:8080
echo      (Open this link in Android Chrome or iPhone Safari on same Wi-Fi)
echo  ----------------------------------------------------------------------
echo.
echo  Press Ctrl+C to stop the server anytime.
echo.
python -m http.server 8080 --bind 0.0.0.0
pause

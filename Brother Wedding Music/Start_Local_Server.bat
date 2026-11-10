@echo off
title Ramesh Wedding Music Server (Local DJ Station)
echo ========================================================
echo   👑 RAMESH WEDDING MUSIC LOCAL HTTP SERVER
echo   10, 11 & 12 December 2026 - 508 Offline Audio Tracks
echo ========================================================
echo.
echo Starting local DJ server on port 8080...
echo You can connect laptops, tablets, and phones on the same Wi-Fi!
echo.
start "" "http://localhost:8080/Wedding_Music_Player.html"
python -m http.server 8080 --directory "%~dp0"
pause

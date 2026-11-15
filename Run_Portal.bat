@echo off
title BDO Bani Integrated Portal
cd /d "%~dp0"
echo =======================================================
echo   Starting BDO Bani Integrated Works Portal & Backend...
echo =======================================================
start "" /B node server.js
timeout /t 2 /nobreak >nul
start "" http://localhost:3000
exit

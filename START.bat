@echo off
:: tambahkan node ke PATH jika belum ada
set PATH=%PATH%;C:\Program Files\nodejs
title MICROSOFT SENDER

:runsender
cls
echo Running index.js...
node index.js
echo.
pause
goto runsender

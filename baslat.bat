@echo off
chcp 65001 >nul
title Dilek Ofis Fiyat Katalogu
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js bulunamadi. Lutfen https://nodejs.org adresinden "LTS" surumunu kurun.
  pause
  exit /b
)
start "" http://localhost:3000/admin
node server.js
pause

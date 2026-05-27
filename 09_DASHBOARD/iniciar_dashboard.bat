@echo off
title Dashboard Roble Fuerte
cd /d "C:\Users\elonc\OneDrive\Documentos\Claude\Projects\ROBLE FUERTE\09_DASHBOARD"

echo.
echo  Iniciando Dashboard Roble Fuerte...
echo.

:: Fechar qualquer servidor antigo na porta 3000
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":3000 "') do (
  taskkill /PID %%a /F >nul 2>&1
)
timeout /t 1 /nobreak >nul

start "" "http://localhost:3000"
"C:\Program Files\nodejs\node.exe" server.js

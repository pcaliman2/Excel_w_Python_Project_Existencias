@echo off
cd /d "%~dp0"

REM variables de entorno
set INVENTARIO_DB_USER=postgres
set INVENTARIO_DB_PASSWORD=...
set INVENTARIO_DB_HOST=nozomi.proxy.rlwy.net
set INVENTARIO_DB_PORT=59582
set INVENTARIO_DB_NAME=railway

start "" /b "%~dp0inventario_4.exe"
exit

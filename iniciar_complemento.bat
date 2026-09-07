@echo off
title Matrix Slide Finder
cd /d "%~dp0"

if not exist "node_modules" (
    echo.
    echo Primera vez usando el complemento: instalando dependencias...
    echo Esto puede tardar 1-2 minutos, es normal.
    echo.
    call npm install
)

echo.
echo Iniciando Matrix Slide Finder y abriendo PowerPoint...
echo (No cierres esta ventana mientras uses el complemento)
echo.
call npm start

pause

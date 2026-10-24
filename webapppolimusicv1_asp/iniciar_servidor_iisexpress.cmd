@echo off
chcp 65001 >nul
title PoliMusic - Servidor de prueba IIS Express

echo ========================================================
echo   PoliMusic - Servidor Local (IIS Express)
echo ========================================================
echo.

set "IIS_EXPRESS=%ProgramFiles%\IIS Express\iisexpress.exe"
set "IIS_EXPRESS_X86=%ProgramFiles(x86)%\IIS Express\iisexpress.exe"
set "PORT=8092"
set "APP_PATH=%~dp0"

if exist "%IIS_EXPRESS%" (
    set "EXE_PATH=%IIS_EXPRESS%"
) else if exist "%IIS_EXPRESS_X86%" (
    set "EXE_PATH=%IIS_EXPRESS_X86%"
) else (
    echo [!] No se encontro IIS Express.
    pause
    exit /b
)

echo Iniciando servidor en el puerto %PORT%...
echo URL: http://localhost:%PORT%/login.asp
echo.
echo Presiona Ctrl+C para detener el servidor.
echo.

start http://localhost:%PORT%/login.asp

"%EXE_PATH%" /path:"%APP_PATH%" /port:%PORT%

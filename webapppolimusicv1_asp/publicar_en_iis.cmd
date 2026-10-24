@echo off
setlocal
REM ------------------------------------------------------------------
REM  Publica PoliMusic en IIS como aplicacion virtual.
REM  Ejecutar como ADMINISTRADOR (clic derecho > Ejecutar como admin)
REM ------------------------------------------------------------------

set APPCMD=%SystemRoot%\System32\inetsrv\appcmd.exe
set APPNAME=PoliMusic
set APPPATH=/polimusic
set PHYS=%~dp0

echo.
echo  ==========================================================
echo   Publicando %APPNAME% en IIS
echo   Ruta fisica: %PHYS%
echo  ==========================================================
echo.

if not exist "%APPCMD%" (
    echo  [ERROR] No se encuentra appcmd.exe. IIS no esta instalado.
    echo.
    pause
    exit /b 1
)

REM --- Permisos NTFS para IIS ---
echo  Asignando permisos a IIS_IUSRS...
icacls "%PHYS%." /grant "IIS_IUSRS":(OI)(CI)RX /t >nul

REM --- Publicar aplicacion virtual en Default Web Site ---
"%APPCMD%" list site "Default Web Site" >nul 2>&1
if %errorlevel%==0 (
    echo  Sitio destino: "Default Web Site"
    "%APPCMD%" delete app /site:"Default Web Site" /path:%APPPATH% >nul 2>&1
    "%APPCMD%" add app /site:"Default Web Site" /path:%APPPATH% /physicalPath:"%PHYS%."
    if errorlevel 1 (
        echo  [ERROR] No se pudo crear la aplicacion virtual.
        pause
        exit /b 1
    )
    set URL=http://localhost%APPPATH%/login.asp
) else (
    echo  Creando sitio independiente en el puerto 8085...
    "%APPCMD%" delete site "%APPNAME%" >nul 2>&1
    "%APPCMD%" create site "%APPNAME%" /physicalPath:"%PHYS%." /bindings:http/*:8085:
    set URL=http://localhost:8085/login.asp
)

echo.
echo  ==========================================================
echo   LISTO. Abre en tu navegador:
echo   %URL%
echo  ==========================================================
echo.
pause
endlocal

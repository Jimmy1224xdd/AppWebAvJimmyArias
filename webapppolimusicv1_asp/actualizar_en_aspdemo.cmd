@echo off
setlocal
echo ========================================================
echo   Copiar PoliMusic a C:\inetpub\wwwroot\aspdemo\
echo   (Ejecutar como ADMINISTRADOR)
echo ========================================================
echo.

set "DEST=C:\inetpub\wwwroot\aspdemo"
set "SRC=%~dp0"

if not exist "%DEST%" (
    mkdir "%DEST%"
)

echo Copiando archivos .asp...
copy /Y "%SRC%*.asp" "%DEST%\"

echo Copiando carpetas css, js y recursos...
xcopy /E /I /Y "%SRC%css" "%DEST%\css" >nul
xcopy /E /I /Y "%SRC%js" "%DEST%\js" >nul
xcopy /E /I /Y "%SRC%SongFilesAndImages_resources" "%DEST%\SongFilesAndImages_resources" >nul

echo Permisos NTFS para IIS_IUSRS...
icacls "%DEST%" /grant "IIS_IUSRS":(OI)(CI)RX /t >nul

if %errorlevel%==0 (
    echo.
    echo ========================================================
    echo  [EXITO] Proyecto completo copiado a aspdemo.
    echo  Abre en tu navegador:
    echo  http://localhost/aspdemo/login.asp
    echo ========================================================
) else (
    echo.
    echo [ERROR] No se pudo copiar. Asegurate de ejecutar este script
    echo como Administrador (clic derecho ^> Ejecutar como administrador).
)
echo.
pause

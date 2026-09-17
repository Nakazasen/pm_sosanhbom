@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
title SSBOM MANAGER - KHOI DONG HE THONG

echo ======================================================================
echo       HE THONG SO SANH BOM TU DONG - SSBOM MANAGER v1.0.0
echo ======================================================================
echo.

rem Kiem tra Python tren he thong
where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo [INFO] Dang khoi chay qua Python Launcher...
    python SSBOM_Launcher.py %*
    goto :done
)

rem Kiem tra ban Portable EXE doc lap
if exist "dist\SSBOM_Portable\SSBOM_Portable.exe" (
    echo [INFO] Dang mo ban Portable EXE doc lap...
    start "" "dist\SSBOM_Portable\SSBOM_Portable.exe" %*
    goto :done
)

echo [LOI] Khong tim thay Python hoac file dist\SSBOM_Portable\SSBOM_Portable.exe!
echo Vui long kiem tra lai moi truong may tinh.
pause

:done
endlocal


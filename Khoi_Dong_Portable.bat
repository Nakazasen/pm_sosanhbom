@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
title SSBOM MANAGER - BAN PORTABLE

echo ======================================================================
echo       SSBOM MANAGER - BAN PORTABLE EXE DOC LAP
echo ======================================================================
echo.

if exist "dist\SSBOM_Portable\SSBOM_Portable.exe" (
    echo [INFO] Dang mo ung dung: dist\SSBOM_Portable\SSBOM_Portable.exe...
    start "" "dist\SSBOM_Portable\SSBOM_Portable.exe" %*
) else (
    echo [LOI] Khong tim thay file dist\SSBOM_Portable\SSBOM_Portable.exe!
    pause
)

endlocal

@echo off
setlocal enabledelayedexpansion
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
    goto :done
)

echo [THONG TIN] Chua co file EXE tai dist\SSBOM_Portable\SSBOM_Portable.exe
echo (Do day la thu muc ma nguon clone tu Git, chua qua buoc build PyInstaller).
echo.
echo [TUDONG] Dang chuyen sang khoi dong bang Python tren may qua Khoi_Dong_SSBOM.bat...
echo.
call "Khoi_Dong_SSBOM.bat" %*

:done
endlocal

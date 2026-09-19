@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
title SSBOM MANAGER - KHOI DONG HE THONG

echo ======================================================================
echo       HE THONG SO SANH BOM TU DONG - SSBOM MANAGER v1.0.0
echo ======================================================================
echo.

rem 1. Kiem tra py launcher (chinh thuc cua Python tren Windows)
py -c "import sys" >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo [INFO] Dang khoi chay qua Windows Python Launcher - py...
    py SSBOM_Launcher.py %*
    goto :done
)

rem 2. Kiem tra cac duong dan Python cu the tren may
if exist "%LocalAppData%\Programs\Python\Python313\python.exe" (
    echo [INFO] Dang khoi chay qua Python 3.13 tai AppData...
    "%LocalAppData%\Programs\Python\Python313\python.exe" SSBOM_Launcher.py %*
    goto :done
)

if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
    echo [INFO] Dang khoi chay qua Python 3.12 tai AppData...
    "%LocalAppData%\Programs\Python\Python312\python.exe" SSBOM_Launcher.py %*
    goto :done
)

if exist "C:\Python313\python.exe" (
    echo [INFO] Dang khoi chay qua C:\Python313...
    "C:\Python313\python.exe" SSBOM_Launcher.py %*
    goto :done
)

if exist "C:\Python312\python.exe" (
    echo [INFO] Dang khoi chay qua C:\Python312...
    "C:\Python312\python.exe" SSBOM_Launcher.py %*
    goto :done
)

rem 3. Kiem tra lenh python tren PATH (loai tru shim WindowsApps bi loi)
python -c "import sys" >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo [INFO] Dang khoi chay qua python trong PATH...
    python SSBOM_Launcher.py %*
    goto :done
)

rem 4. Kiem tra ban Portable EXE doc lap
if exist "dist\SSBOM_Portable\SSBOM_Portable.exe" (
    echo [INFO] Dang mo ban Portable EXE doc lap...
    start "" "dist\SSBOM_Portable\SSBOM_Portable.exe" %*
    goto :done
)

echo.
echo [LOI] Khong the tim thay moi truong Python hop le hoac file Portable EXE!
echo ======================================================================
echo Huong dan khac phuc:
echo 1. Cai dat Python 3.11 - 3.13 tu https://www.python.org (tich vao 'Add Python to PATH')
echo 2. Hoac build ban Portable EXE: chay lenh 'py -m PyInstaller pm_sosanhbom.spec'
echo ======================================================================
pause

:done
endlocal

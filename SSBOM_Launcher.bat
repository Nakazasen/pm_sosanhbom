@echo off
rem Launcher Batch Script for SSBOM Manager
setlocal
cd /d "%~dp0"
py SSBOM_Launcher.py %*
endlocal

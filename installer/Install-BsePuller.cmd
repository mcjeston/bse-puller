@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Install-BsePuller.ps1"
exit /b %errorlevel%

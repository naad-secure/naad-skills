@echo off
rem word-live on Windows: screenshot of this screen. Usage: shot.cmd [out.png]
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0shot.ps1" %*
exit /b %ERRORLEVEL%

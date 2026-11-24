@echo off
rem word-live on Windows: run a PowerShell snippet in the Word open on this PC.
rem Usage: word.cmd job.ps1   (the snippet saved as a UTF-8 file; stdin also works)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0word.ps1" %*
exit /b %ERRORLEVEL%

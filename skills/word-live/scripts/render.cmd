@echo off
rem word-live on Windows: like word.cmd; each Render N in the job writes PREFIX-1.png, PREFIX-2.png ...
rem Usage: render.cmd PREFIX job.ps1
if "%~1"=="" (echo usage: render.cmd ^<prefix^> job.ps1 1>&2 & exit /b 2)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0word.ps1" -RenderPrefix "%~1" %2 %3
exit /b %ERRORLEVEL%

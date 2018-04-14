@echo off
echo Running from cmd.exe. Consider using Powershell.exe directly for faster startup
set pwsh=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
%pwsh% -NoProfile -ExecutionPolicy Bypass -File %~dp0%~n0.ps1 -FromBatch -PathExt "%PathExt%" %*
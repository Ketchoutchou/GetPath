@echo off
echo Running from cmd.exe. Consider using Powershell.exe directly for faster startup
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File %~dp0GetPath.ps1 -FromBatch %*
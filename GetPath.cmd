@echo off
echo Running from cmd.exe. Consider using Powershell.exe directly for faster startup
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe %~dp0GetPath.ps1 -FromBatch %*
REM Will probably need https://blog.netspi.com/15-ways-to-bypass-the-powershell-execution-policy/
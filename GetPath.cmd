@echo off
echo Running from cmd.exe. Consider using Powershell.exe directly for faster startup.
powershell %~dp0GetPath.ps1 %*
REM Will probably need https://blog.netspi.com/15-ways-to-bypass-the-powershell-execution-policy/
@echo off
echo Running from cmd.exe. Consider using Powershell.exe directly for faster startup
if exist "%ProgramW6432%\PowerShell\pwsh.exe" (
	set pwsh=C:\Program Files\PowerShell\pwsh.exe
) else (
	set pwsh=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
)
"%pwsh%" -NoProfile -ExecutionPolicy Bypass -File %~dp0%~n0.ps1 -FromBatch -PathExt "%PathExt%" %*
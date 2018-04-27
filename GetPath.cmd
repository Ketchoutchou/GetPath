@echo off
setlocal
echo Consider using PowerShell directly for faster startup

if "%*"=="" (
	goto :RunWithPowerShell
)

set params=%*
if not "x%params:-Reload=%"=="x%params%" (
	goto :ReloadPath
) else (
	goto :RunWithPowerShell
)

:ReloadPath
call :SetFromReg "HKLM\System\CurrentControlSet\Control\Session Manager\Environment" Path Path_HKLM
call :SetFromReg "HKCU\Environment" Path Path_HKCU

if "%PATH_HKLM:~-1%"==";" (
	set "sep="
) else (
	set "sep=;"
)

if "%PATH_HKLM%"=="" set "sep="
if "%PATH_HKCU%"=="" set "sep="

REM call is used to expand variable inside another variable
REM we need to endlocal in order to set PATH globally (for this console)
endlocal && call set "PATH=%PATH_HKLM%%sep%%PATH_HKCU%" && setlocal

C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -Command ^
$colorBefore = $host.ui.RawUI.BackgroundColor; ^
$consoleId = (Get-WmiObject Win32_Process -Filter ProcessId=$PID).ParentProcessId; ^
$console = Get-Process -Id $consoleId; ^
$host.ui.RawUI.BackgroundColor = """DarkMagenta"""; ^
echo """Path environment variable for $($console.MainModule.ModuleName) (PID:$consoleId) has been reloaded from registry"""; ^
$host.ui.RawUI.BackgroundColor = $colorBefore

:RunWithPowerShell

if exist "%ProgramW6432%\PowerShell\pwsh.exe" (
	set getpath_pwsh=C:\Program Files\PowerShell\pwsh.exe
) else (
	set getpath_pwsh=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
)
if defined GetPath_PowerShell (
	set getpath_pwsh=%GetPath_PowerShell%
)

call "%getpath_pwsh%" -NoProfile -ExecutionPolicy Bypass -File %~dp0%~n0.ps1 -FromBatch -PathExt "%PathExt%" %*
exit /b %errorlevel%

:SetFromReg
	REM We use full path for reg.exe in case PATH is broken or empty
	REM reg.exe error output is redirected to NUL in case key doesn't exist
	for /f "usebackq skip=2 tokens=2,*" %%A in (`C:\Windows\System32\reg.exe query "%~1" /v "%~2" 2^>nul`) do (
		set "%~3=%%B"
	)
	goto :eof
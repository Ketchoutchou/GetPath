@echo off
echo Running from cmd.exe. Consider using Powershell.exe directly for faster startup

if exist "%ProgramW6432%\PowerShell\pwsh.exe" (
	set getpath_pwsh=C:\Program Files\PowerShell\pwsh.exe
) else (
	set getpath_pwsh=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
)
if defined PowerShell (
	set getpath_pwsh=%PowerShell%
)

if "%*"=="" (
	goto :RunWithPowerShell
)
set params=%*
if not "x%params:-Reload=%"=="x%params%" goto :ReloadPathAndRun
goto :RunWithPowerShell

:ReloadPathAndRun
	REM setlocal cannot be use here because we are setting PATH environment variable for the current session
	REM Hence, used variables need to be stored
	REM Evert sey call are protected with quotes against spaces
	set "OLD_PATH_HKLM=%PATH_HKLM%"
	set "OLD_PATH_HKCU=%PATH_HKCU%"
	set "PATH_HKLM="
	set "PATH_HKCU="

	call :SetFromReg "HKLM\System\CurrentControlSet\Control\Session Manager\Environment" Path Path_HKLM
	call :SetFromReg "HKCU\Environment" Path Path_HKCU

	if "%PATH_HKLM:~-1%"==";" (
		set "SEP="
	) else (
		set "SEP=;"
	)

	if "%PATH_HKLM%"=="" set "SEP="
	if "%PATH_HKCU%"=="" set "SEP="

	REM call is used to expand variable inside another variable
	call set "PATH=%PATH_HKLM%%SEP%%PATH_HKCU%"
	echo Path environment variable (for cmd.exe) has been reloaded from registry

	REM endlocal cannot be used here because we are setting PATH environment variable for the current session
	REM Hence, used variables need to be restored
	set "SEP="
	set "PATH_HKLM=%OLD_PATH_HKLM%"
	set "PATH_HKCU=%OLD_PATH_HKCU%"
	set "OLD_PATH_HKLM="
	set "OLD_PATH_HKCU="

	goto :RunWithPowerShell

:SetFromReg
	REM We use full path for reg.exe in case PATH is broken or empty
	REM reg.exe error output is redirected to NUL in case key doesn't exist
	for /f "usebackq skip=2 tokens=2,*" %%A in (`C:\Windows\System32\reg.exe query "%~1" /v "%~2" 2^>nul`) do (
		set "%~3=%%B"
	)
	goto :eof
	
:RunWithPowerShell
	"%getpath_pwsh%" -NoProfile -ExecutionPolicy Bypass -File %~dp0%~n0.ps1 -FromBatch -PathExt "%PathExt%" %*
	exit /b %errorlevel%
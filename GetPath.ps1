<#
.SYNOPSIS
GetPath helps you detect and fix issues in your PATH environment variable on Windows

.DESCRIPTION
and so much more

.INPUTS
Process name, id or object

.OUTPUTS
System.Int. GetPath returns a simple exit code based on the result of the analysis

.EXAMPLE
C:\PS> GetPath -DontCheckUnexpandedDuplicates
Current PATH environment variable is 138 character long (maximum is 2047)
---------- PATH BEGIN ----------
%SystemRoot%\system32
%SystemRoot%
%SystemRoot%\System32\Wbem
%SYSTEMROOT%\System32\WindowsPowerShell\v1.0
C:\ProgramData\Oracle\Java\javapath
----------- PATH END -----------

.LINK
https://github.com/Ketchoutchou/GetPath
#>

[CmdletBinding()]
Param(
	# Add the specified entry to the selected PATH environment variable.
	# Requires -ToUserPath or -ToSystemPath.
	# Supports short names and unexpanded variables.
	[Alias("A")]
	[string] $AddEntry = "",
	
	# Do not take variable-based path entries into account.
	[switch] $DontCheckUnexpandedDuplicates = $false,

	# Internal parameter to know if GetPath has been launched using GetPath.cmd.
	[Parameter( <# DontShow #> )] #Need to restrict script to PowerShell >=5
	[switch] $LaunchedFromBatch = $false,
	
	# Analyze PATH from registry, ignoring current context modification.
	[Alias("Registry")]
	[switch] $FromRegistry = $false,
	
	# Internal parameter (used if GetPath has been launched using GetPath.cmd) to retrieve PathExt environment variable value from cmd.exe.
	[Parameter( <# DontShow #> )] #Need to restrict script to PowerShell >=5
	[string] $PathExt = "",
	
	# Analyze PATH from string parameter
	# Can be set from pipeline
	[Parameter(ValueFromPipeline = $true)]
	[Alias("String")]
	[string] $FromString = "",

	# Get PATH environment variable from another running process in real time (using process id or approximate name).
	[Parameter(Position = 0)]
	[Alias("ProcessNameOrId", "P")]
	[string] $FromProcessNameOrId = "",
	
	# Get PATH environment variable from another running process in real time (using process object).
	# Can be set from pipeline
	[Parameter(ValueFromPipeline = $true)]
	[Alias("ProcessObject")]
	[System.Diagnostics.Process] $FromProcessObject,
	
	# Replace current context PATH environment variable with the one found in registry.
	# If launched using GetPath.cmd, only -Reload is supported
	[Alias("Refresh", "R")]
	[switch] $Reload = $false,
	
	# Remove the specified entry from the selected PATH environment variable.
	# Requires -ToUserPath or -ToSystemPath.
	# Supports short names and unexpanded variables.
	[Alias("DeleteEntry")]
	[string] $RemoveEntry = "",
	
	# Analyze only the system PATH environment variable.
	# If -AddEntry or -RemoveEntry is set, it will add or remove the entry to the system PATH environment variable (requires administrative rights).
	[Alias("System", "ToMachinePath", "Machine")]
	[switch] $ToSystemPath = $false,
	
	# Internal parameter used for testing.
	[Parameter( <# DontShow #>)] #Need to restrict script to PowerShell >=5
	[switch] $TestMode = $false,

	# Analyze only the user PATH environment variable.
	# If -AddEntry or -RemoveEntry is set, it will add or remove the entry to the user PATH environment variable.
	[Alias("User")]
	[switch] $ToUserPath = $false,
	
	# Remove any prefix when displaying PATH entries.
	# Good for copy/pasting.
	[Alias("Porcelain")]
	[switch] $Verbatim = $false,
	
	# Show the current version number
	[Alias("About")]
	[switch] $Version = $false,
	
	# Find all occurrences of an executable in the current context PATH.
	# This parameter supports bulk search using wildcards.
	[Alias("Which", "Search", "W")]
	[string] $Where = ""
)

Set-StrictMode -Version Latest
#$ErrorActionPreference = "Stop"

if ($DebugPreference -eq "Inquire") {
	$DebugPreference = "Continue"
	Write-Debug "Switching DebugPreference to Continue"
}

<#  Need to restrict script to PowerShell >=5
Class Chrono {
	[String] $Comment
	[DateTime] $Chrono
	[Int] $ShowLongerThan
	
	Chrono ([String] $Comment, [Int] $ShowLongerThan) {
		$this.Comment = $Comment
		$this.ShowLongerThan = $ShowLongerThan
		$this.Chrono = Get-Date
		Write-Verbose "'$($this.Comment)' started on $($this.Chrono)"
	}
	
	Stop() {
		$duration = $(Get-Date) - $this.Chrono
		$totalTime = $duration.TotalMilliseconds
		if ($totalTime -gt $this.ShowLongerThan) {
			Write-Debug "'$($this.Comment)' took $($totalTime.ToString('#')) ms"
		}
	}
}
#>

function ShowVersion {
	if ($Version) {
		echo @'

 $$$$$$\             $$\     $$$$$$$\            $$\     $$\       
$$  __$$\            $$ |    $$  __$$\           $$ |    $$ |      
$$ /  \__| $$$$$$\ $$$$$$\   $$ |  $$ |$$$$$$\ $$$$$$\   $$$$$$$\  
$$ |$$$$\ $$  __$$\\_$$  _|  $$$$$$$  |\____$$\\_$$  _|  $$  __$$\ 
$$ |\_$$ |$$$$$$$$ | $$ |    $$  ____/ $$$$$$$ | $$ |    $$ |  $$ |
$$ |  $$ |$$   ____| $$ |$$\ $$ |     $$  __$$ | $$ |$$\ $$ |  $$ |
\$$$$$$  |\$$$$$$$\  \$$$$  |$$ |     \$$$$$$$ | \$$$$  |$$ |  $$ |
 \______/  \_______|  \____/ \__|      \_______|  \____/ \__|  \__| 2.5

'@
		exit 1408
	}
}
function PSVersionCheck {
	$majorVersion = $PSVersionTable.PSVersion.Major
	$minorVersion = $PSVersionTable.PSVersion.Minor
	if($majorVersion -lt 5){
		Write-Warning "You are using PowerShell $majorVersion.$minorVersion. Consider upgrading to PowerShell 5.0 at least for better performance."
	}
}
function GetPathFromRegistry { 
	Param (
		[parameter(Mandatory=$true)] [String]$regPath
	)
	
	(Get-Item $regPath).GetValue('PATH',$null,[Microsoft.WIN32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
}
function JoinSystemAndUserPath {
	Param (
		[String]$systemPath,
		$userPath
	)
	
	if (!$systemPath.EndsWith(';') -And $systemPath -And $userPath -is [String]) {
		$separator = ";"
	} else {
		$separator = ""
	}
	$systemPath + "$separator" + $userPath
}
function ShowPathLength { 
	Param (
		[String]$path
	)
	
	$warningLength = 1600
	$expandedPathLength = [System.Environment]::ExpandEnvironmentVariables($path).Length
	if ($expandedPathLength -gt 2047 ) {
		$host.ui.RawUI.ForegroundColor = "Red"
		Write-Warning "See https://software.intel.com/en-us/articles/limitation-to-the-length-of-the-system-path-variable"
	} elseif (($expandedPathLength -eq 0)){
		$host.ui.RawUI.ForegroundColor = "Red"
	} elseif ($expandedPathLength -gt $warningLength) {
		$host.ui.RawUI.ForegroundColor = "Yellow"
	} else {
		$host.ui.RawUI.ForegroundColor = "Green"
	}
	echo "Current PATH environment variable is $expandedPathLength character long (maximum is 2047)"
	$host.ui.RawUI.ForegroundColor = "Gray"
	if ($expandedPathLength -gt 4095) {
		Write-Error "Path too long. Not implemented yet"
		exit -1
	}
}
function PrepareForTestPath {
	Param (
		[String]$pathEntry
	)
	
 	$pathEntry = $pathEntry.Replace('"','')
	$pathEntry = $pathEntry.Replace("`t",'')
	$pathEntry = $pathEntry.Trim()
	if (!$DontCheckUnexpandedDuplicates){
		$pathEntry = [System.Environment]::ExpandEnvironmentVariables($pathEntry)
	}
	$pathEntry
}
function NewPrepareForTestPath {
	Param (
		[String]$pathEntry
	)
	
	$pathEntry = [System.Environment]::ExpandEnvironmentVariables($pathEntry)
	$pathEntry
}
function GetShortPathEntry {
	Param (
		[String]$pathEntry
	)

	if ($pathEntry) {
		$pathEntry = PrepareForTestPath($pathEntry)
		try {
			if ($pathEntry -And (Test-Path $pathEntry -IsValid) -And (Test-Path $pathEntry -ErrorAction Stop)) {
				$FSO = New-Object -ComObject Scripting.FileSystemObject
				return $FSO.GetFolder($pathEntry).ShortPath
			} else {
				return $pathEntry.TrimEnd('\')
			}
		} catch {
			if ($_Exception.Message -eq "Access is denied") {
				return $pathEntry.TrimEnd('\')
			} else {
				throw
			}
		}
	} else {
		return $null
	}
}
function GetFullPathEntry {	
	Param (
		[String]$pathEntry
	)

	if ($pathEntry) {
		$pathEntry = PrepareForTestPath($pathEntry)
		if ($pathEntry -And (Test-Path $pathEntry -IsValid) -And (Test-Path $pathEntry)) {
			return (Get-Item $pathEntry -Force).FullName
		} else {
			return $pathEntry.TrimEnd('\')
		}
	} else {
		return $null
	}
}
function DisplayPathEntry {
	Param (
		[String]$pathEntry
	)

	$displayPathBeginChar = '"'
	$displayPathEndChar = '"'
	"$displayPathBeginChar" + "$pathEntry" + "$displayPathEndChar"
}
function DisplayPathEntryWithOrder {
	Param (
		[HashTable]$pathEntry
	)

	$displayPathEntry = DisplayPathEntry $pathEntry.OriginalPath
	$displayPathEntryOrder = $pathEntry.EntryOrder
	"$displayPathEntry (Entry #$displayPathEntryOrder)"
}
function ListDuplicates { 
	Param (
		[parameter(Mandatory=$true)] [System.Collections.ArrayList]$pathChecker
	)
	
	if($PSVersionTable.PSVersion.Major -gt 2) {
		$pristinePathsList = $pathChecker.PristinePath
	} else {
		$pristinePathsList = foreach ($pathCheckerEntry in $pathChecker) {$pathCheckerEntry.PristinePath}
	}
	$duplicates = [array]$pristinePathsList | Group | ? { $_.Count -gt 1 }
	if ($duplicates) {
		$duplicateCount = @($duplicates).Length
		$host.ui.RawUI.ForegroundColor = "Cyan"
		echo "`n____________________________________________"
		echo "Duplicated path entries have been found ($duplicateCount):`n"
		foreach ($duplicate in $duplicates) {
			$fullPathEntry = GetFullPathEntry($duplicate.Name)
			$host.ui.RawUI.ForegroundColor = "Cyan"
			echo " $(DisplayPathEntry $fullPathEntry) ($($duplicate.Count) occurrences):"
			$host.ui.RawUI.ForegroundColor = "Gray"
			foreach ($pathCheckerEntry in $pathChecker) {
				if ($pathCheckerEntry.PristinePath -eq $duplicate.Name) {
					echo "    ->  $(DisplayPathEntryWithOrder $pathCheckerEntry)"
				}
			}
		}
		$host.ui.RawUI.ForegroundColor = "Gray"
	}
}
function ListIssues {
	Param (
		[parameter(Mandatory=$true)] [System.Collections.ArrayList]$pathChecker
	)
	
	foreach ($pathCheckerEntry in $pathChecker) {
		if (!$pathCheckerEntry.OriginalPath){
			$pathCheckerEntry.Issues.Add("MustNotBeEmpty") | Out-Null
			continue
		}
		if ($pathCheckerEntry.OriginalPath -match '^[ ]+$') {
			$pathCheckerEntry.Issues.Add("MustNotContainOnlySpaces") | Out-Null
			continue
		}
		if ($pathCheckerEntry.OriginalPath.Contains('"')){
			$pathCheckerEntry.Issues.Add("ShouldNotContainQuotes") | Out-Null
		}
		if ($pathCheckerEntry.OriginalPath.Contains('*')){
			$pathCheckerEntry.Issues.Add("MustNotContainAsterisk") | Out-Null
		}
		if ($pathCheckerEntry.OriginalPath.Contains('>') -Or $pathCheckerEntry.OriginalPath.Contains('<')){
			$pathCheckerEntry.Issues.Add("MustNotContainAngleBracket") | Out-Null
		}
		if ($pathCheckerEntry.OriginalPath.Contains('?')){
			$pathCheckerEntry.Issues.Add("MustNotContainQuestionMark") | Out-Null
		}
		if ($pathCheckerEntry.OriginalPath.Contains('|')){
			$pathCheckerEntry.Issues.Add("MustNotContainPipe") | Out-Null
		}
		if ($pathCheckerEntry.OriginalPath.TrimStart("`t").Contains("`t")){
			$pathCheckerEntry.Issues.Add("MustNotContainTabs") | Out-Null
		}
		if ($pathCheckerEntry.OriginalPath.StartsWith("`t")){
			$pathCheckerEntry.Issues.Add("ShouldNotStartWithTabs") | Out-Null
		}
		if ($pathCheckerEntry.OriginalPath.StartsWith(' ')){
			$pathCheckerEntry.Issues.Add("ShouldNotStartWithSpace") | Out-Null 
		}
		if ($pathCheckerEntry.OriginalPath.EndsWith(' ')){
			$pathCheckerEntry.Issues.Add("MustNotEndWithSpace") | Out-Null
		}
		if ($pathCheckerEntry.OriginalPath.EndsWith('\')){
			$pathCheckerEntry.Issues.Add("ShouldNotEndWithBackslash") | Out-Null
		}
		try {
			if (!(Test-Path $pathCheckerEntry.OriginalPath -IsValid)){
				$pathCheckerEntry.Issues.Add("ShouldBeValid") | Out-Null
			} elseif (!(Test-Path (NewPrepareForTestPath($pathCheckerEntry.OriginalPath)) -ErrorAction Stop)){
				$pathCheckerEntry.Issues.Add("MustExist") | Out-Null
			}
		} catch {
			if ($_Exception.Message -eq "Access is denied") {
				$pathCheckerEntry.Issues.Add("AccessDenied") | Out-Null
			} else {
				throw
			}
		}
	}
	[Array]$issuesList = foreach ($pathCheckerEntry in $pathChecker) {$pathCheckerEntry.Issues}
	if ($issuesList) {
		$host.ui.RawUI.ForegroundColor = "Yellow"
		echo "`n___________________________"
		echo "Issues have been found ($(@($issuesList).Count)):`n"
		foreach ($pathCheckerEntry in $pathChecker) {
			if ($pathCheckerEntry.Issues.Count -gt 0){
				$originalPath = $pathCheckerEntry.OriginalPath
				$host.ui.RawUI.ForegroundColor = "Yellow"
				echo " $(DisplayPathEntryWithOrder $pathCheckerEntry)"
				$host.ui.RawUI.ForegroundColor = "Gray"
				foreach ($warning in $pathCheckerEntry.Issues) {
					echo "    ->  $warning"
				}
			}
		}
		$host.ui.RawUI.ForegroundColor = "Gray"
	}
}
function GetPathPrefix {
	Param (
		[parameter(Mandatory=$true)] $pathEntry
	)
	
	$flags = "$($pathEntry.EntryOrder)`t"
	if ($pathEntry.IsNetworkPath) {
		$flags += "n"
	} else {
		$flags += "-"
	}
	if ($pathEntry.UnexpandedEntry) {
		$flags += "%"
	} else {
		$flags += "-"
	}
	if ($pathEntry.OriginalPath -like "*~*") {
		$flags += "8"
	} else {
		$flags += "-"
	}
	$flags += "`t  "
	$flags
}
function GetWhereResults {
	Param (
		[parameter(Mandatory=$true)] $pathEntry
	)

	$foundFileList = @()
	$searchPattern = $pathEntry.PristinePath
	#$chrono = [Chrono]::new("GCI", 20)
	if($PSVersionTable.PSVersion.Major -gt 2) {
		$fileList = gci -Force -File $searchPattern -Filter $filter -ErrorAction SilentlyContinue
	} else {
		$fileList = gci -Force $searchPattern -Filter $filter -ErrorAction SilentlyContinue | where { $_.GetType().Name -eq "FileInfo" }
	}
	#$chrono.Stop()

	if ($filelist) {
		if (!$containsWildcard) {
			#$chrono = [Chrono]::new("Where", 20)
			if ($containsDot) {
				foreach ($file in $fileList) {
					if ($file.Name -like $where) {
						$foundFileList += $file
						continue
					}
				}
			}
			if (!$LaunchedFromBatch) {
				foreach ($file in $fileList) {
					if ($file.Name -like "$where.ps1") {
						$foundFileList += $file
						continue
					}
				}
			}
			foreach ($pathExtEntry in $pathExtEntries) {
				foreach ($file in $fileList) {
					if ($file.BaseName -like $where -And $file.Extension -eq $pathExtEntry) {
						$foundFileList += $file
						continue
					}
				}
			}
			if (!$LaunchedFromBatch) {
				if (!$containsDot) {
					foreach ($file in $fileList) {
						if ($file.Name -like $where) {
							$foundFileList += $file
							continue
						}
					}
				}
			}
			#$chrono.Stop()
		} else {
			if ($pathEntry.IsNetworkPath) {
				#$chrono = [Chrono]::new("Sort", 20)
				$foundFileList = $fileList | Sort
				#$chrono.Stop()
			} else {
				$foundFileList = $fileList
			}
		}
	}
	$foundFileList
}
function DisplayWhereResults {
	#retrieve color before ?
	$host.ui.RawUI.ForegroundColor = "DarkCyan"
	if (!$Verbatim -And $where -ne "" -And $foundFileList) {
		if ($containsWildcard) {
			if ($foundFileList -is [array]) {
				$fileCount = $foundFileList.Length
			} else {
				$fileCount = 1
			}
			echo "`t`t`t$fileCount file(s) found:"
			if ($PSVersionTable.PSVersion.Major -gt 5) {
				$foundFileList | Format-Wide -AutoSize -Property Name
			} else {
				$($foundFileList | Format-Wide -AutoSize -Property Name | Out-String).Trim()
			}
		} else {
			foreach ($foundFile in $foundFileList) {
				if (!$exactWhereFound) {
					$host.ui.RawUI.ForegroundColor = "Cyan"
				}
				echo "`t`t`t$($foundFile.Name)"
				if (!$exactWhereFound) {
					$host.ui.RawUI.ForegroundColor = "DarkCyan"
					$script:exactWhereFound = $true
				}
			}
		}
	}
	$host.ui.RawUI.ForegroundColor = $colorBefore
}
function DisplayPath {
	Param (
		[parameter(Mandatory=$true)] [System.Collections.ArrayList]$pathChecker,
		[bool]$diffMode,
		[string]$where
	)
	
	$i = 0
	$registryPathEntries = [System.Environment]::ExpandEnvironmentVariables($registryPathString).Split(';')
	$registryPathEntriesCount = $registryPathEntries.Length
	
	if ($where -ne "") {
		if ($LaunchedFromBatch) {
			$pathExtEntries = $PathExt.Split(';')
		} else {
			$pathExtEntries = $env:PathExt.Split(';')
		}
		$containsDot = $where -like "*.*"
		$containsWildcard = $where -match "\*"
		if ($containsWildcard) {
			$filter = $where
		} else {
			$filter = "$where*"
		}
		$script:exactWhereFound = $false
		if ($LaunchedFromBatch) {
			$currentDirEntry = @{
				PristinePath = $pwd
				IsNetworkPath = $false #to fix
				# need more ?
			}
			$foundFileList = GetWhereResults $currentDirEntry
			if ($foundFileList -And !$Verbatim) {
				$colorBefore = $host.ui.RawUI.ForegroundColor
				$host.ui.RawUI.ForegroundColor = "Magenta"
				echo "`t`t  $pwd (current directory; not in your actual PATH)"
				DisplayWhereResults
				$host.ui.RawUI.ForegroundColor = $colorBefore
			}
		}
	}
	foreach ($pathCheckerEntry in $pathChecker) {
		$colorBefore = $host.ui.RawUI.ForegroundColor
		if (!$Verbatim) {
			$prefix = GetPathPrefix $pathCheckerEntry
		} else {
			$prefix = $null
		}
		
		if (!$Verbatim -And $where -ne "") {
			$foundFileList = GetWhereResults $pathCheckerEntry
			if ($foundFileList) {
				$host.ui.RawUI.ForegroundColor = "Magenta"
			}
		}
		
		if ($Verbatim -Or !$diffMode) {
			if ($pathCheckerEntry.OriginalPath -eq "") {
				echo "$prefix<empty>"
			} else {
				echo "$prefix$($pathCheckerEntry.OriginalPath)"
			}
			$i = $registryPathEntriesCount
		} else {
			if ($i -lt $registryPathEntriesCount -And $pathCheckerEntry.OriginalPath -eq $registryPathEntries[$i]) {
				echo "$prefix$($pathCheckerEntry.OriginalPath)"
				$i++
			} else {
				if($PSVersionTable.PSVersion.Major -gt 2) {
					$indexInRegistry = $registryPathEntries.IndexOf($pathCheckerEntry.OriginalPath)
				} else {
					$indexInRegistry = [array]::indexof($registryPathEntries, $pathCheckerEntry.OriginalPath)
				}
				if ($indexInRegistry -ne -1 -And $indexInRegistry -gt $i) {
					for ($j = $i; $j -lt $indexInRegistry; $j++) {
						$host.ui.RawUI.ForegroundColor = "Red"
						echo "`t`t- $($registryPathEntries[$j]) (not present in this context; only in registry)"
					}
					$i = $indexInRegistry + 1
					$host.ui.RawUI.ForegroundColor = $colorBefore
					echo "$prefix$($pathCheckerEntry.OriginalPath)"
				} else {
					$host.ui.RawUI.ForegroundColor = "Yellow"
					echo "$($prefix.TrimEnd("  "))+ $($pathCheckerEntry.OriginalPath) (only present in this context; not in registry)"
				}
			}
		}
		DisplayWhereResults
	}
	if ($i -lt $registryPathEntriesCount) {
		for ($j = $i; $j -lt $registryPathEntriesCount; $j++) {
			$host.ui.RawUI.ForegroundColor = "Red"
			if ($registryPathEntries[$j] -eq "") {
				echo "`t`t- <empty> (not present in this context; only in registry)"
			} else {
				echo "`t`t- $($registryPathEntries[$j]) (not present in this context; only in registry)"
			}
		}
		$host.ui.RawUI.ForegroundColor = $colorBefore
	}
}
function OpenProcessExplorerOffer {
	[string]$processFinder = Read-Host "Need help? Type 'procexp' or 'pslist'"
	if ($processFinder -eq "pslist") {
		& pslist -t -accepteula
	}
	if ($processFinder -eq "procexp") {
		& procexp -accepteula
	}
}
function Main {
	ShowVersion
	PSVersionCheck
	if ($RemoveEntry) {
		$FromRegistry = $true
		echo "Pretending to remove $RemoveEntry"
	}
	if ($AddEntry) {
		$FromRegistry = $true
		echo "Pretending to add $AddEntry"
	}
	if (($AddEntry -Or $RemoveEntry) -And $LaunchedFromBatch -And $Reload) {
		Write-Warning "Cannot reload PATH after adding/removing entries when launched using GetPath.cmd"
	}
	$systemRegistryPathString = GetPathFromRegistry "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
	$userRegistryPathString = GetPathFromRegistry "HKCU:\Environment"
	
	if ($FromProcessObject) {
		$FromProcessNameOrId = $FromProcessObject.Id
	}
	
	if (!$FromRegistry -And $FromProcessNameOrId -ne "") {
		$getExternalProcessPathExecutable = "GetExternalProcessEnv.exe"
		if (Test-Path $scriptRoot\$getExternalProcessPathExecutable) {
			$externalProcessPathString = $null
			$exitCode = -666
			if ($FromProcessNameOrId -match '^\d+$') {
				$process = Get-Process -Id $FromProcessNameOrId
				Write-Warning "Analyzing process $($process.Name) (PID $($process.Id))"
				$externalProcessPathString = & $scriptRoot\$getExternalProcessPathExecutable $FromProcessNameOrId
				$exitCode = $LASTEXITCODE
			}
			if ($exitCode -eq -4) {
				Write-Warning "Access to this process is denied. Please run as administrator."
				exit -4
			}
			if ($exitCode -eq 0) {
				$actualPathString = $externalProcessPathString
			} else {
				try {
					$foundProcesses = Get-Process $FromProcessNameOrId -ErrorAction 'Stop'
				} catch {
					$processList = Get-Process
					$foundProcesses = @()
					foreach ($process in $processList) {
						if ($process.Name -like "*$FromProcessNameOrId*") {
							$foundProcesses += $process
						}
					}
				}
				if ($foundProcesses -is [array]) {
					if ($foundProcesses.Length -gt 1) {
						Write-Warning "Multiple processes found with name containing '$FromProcessNameOrId' (most recent first):"
						$foundProcesses | Add-Member -MemberType NoteProperty -Name TryStartTime -Value " ACCESS DENIED"
						foreach ($foundProcess in $foundProcesses) {
							if ($foundProcess.StartTime) {
								$foundProcess.TryStartTime = $foundProcess.StartTime
							}
						}
						$foundProcesses | Sort TryStartTime -Desc | Format-Table Id, Name, TryStartTime, MainWindowTitle
						do {
							try {
								$InputOK = $true
								[int]$processId = Read-Host "Choose a process id"
								if ($processId -eq "") {
									$InputOK = $false
									OpenProcessExplorerOffer
								}
							} catch {
								$InputOK = $false
								OpenProcessExplorerOffer
							}
						}
						while (!$InputOK)
						# will do better
						& $scriptRoot\GetPath.ps1 -FromProcessNameOrId $processId
						exit $LASTEXITCODE
					}
					if ($foundProcesses.Length -eq 0) {
						Write-Warning "No process found with name containing '$FromProcessNameOrId'"
						exit -3
					}
					$process = $foundProcesses[0]
				} else {
					$process = $foundProcesses
				}
				if (!$process) {
					Write-Warning "No process found with name containing '$FromProcessNameOrId'"
					exit -3
				}
				Write-Warning "Analyzing process $($process.Name) (PID $($process.Id))"
				$externalProcessPathString = & $scriptRoot\$getExternalProcessPathExecutable $process.Id
				if ($LASTEXITCODE -eq -4) {
					Write-Warning "Access to this process is denied. Please run as administrator."
					exit -4
				}
				$actualPathString = $externalProcessPathString
			}
		} else {
			Write-Warning "$getExternalProcessPathExecutable not found. Cannot get PATH of an external process."
			exit -2
		}
	}
	
	if ($FromString -And !$FromProcessObject) {
		Write-Warning "Parameter-only analysis. Current context is ignored"
		$actualPathString = $FromString
	}
	
	if ($TestMode) {
		$systemRegistryPathString = @'
		C:\windows;;s:\;::;\\;s;s:;C:\WINDOWS; C:\windows\;C:/windows/;c:\>;c:;c;c:\windows\\;c:\fdsf\\;\\\;c:\<;%USERPROFILE%\desktop;"c:\program files (x86)"\google;"c:\program files (x86)"\google2;  ;C:\Users\Ketchoutchou\Desktop;c:\doesnotexist;c:\dOesnotexist\ ; c:\program files;C:\windows*;*;?;|;c:|windows;c:\windows?;c:\program files (x86);%SystemRoot%\system32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SYSTEMROOT%\System32\WINDOW~1\v1.0;C:\ProgramData\Oracle\Java\javapath;C:\Program Files (x86)\NVIDIA Corporation\PhysX\Common;D:\Logiciels\Utilitaires\InPath;"c:\program files (x86)"; 
'@
		$userRegistryPathString = @'
C:\userpath
'@
		$actualPathString = [System.Environment]::ExpandEnvironmentVariables($(JoinSystemAndUserPath $systemRegistryPathString $userRegistryPathString))
	}
	
	$registryPathString = JoinSystemAndUserPath $systemRegistryPathString $userRegistryPathString
	$expandedRegistryPathString = [System.Environment]::ExpandEnvironmentVariables($registryPathString)
	if (!$LaunchedFromBatch -And $Reload) {
		$env:PATH = $expandedRegistryPathString
		$colorBefore = $host.ui.RawUI.BackgroundColor
		$host.ui.RawUI.BackgroundColor = "DarkMagenta"
		$console = Get-Process -Id $pid
		echo "Path environment variable for $($console.MainModule.ModuleName) (PID:$pid) has been reloaded from registry"
		$host.ui.RawUI.BackgroundColor = $colorBefore
	}
	if (!$(Test-Path variable:actualPathString)) {
		$actualPathString = $env:PATH
	}
	
	$diffMode = $false
	if ($FromRegistry -Or $expandedRegistryPathString -eq $actualPathString) {
		if ($FromRegistry) {
			Write-Warning "Registry-only analysis. Current context is ignored."
		}
		if ($userRegistryPathString -eq "") {
			Write-Warning "User PATH environment variable is defined but empty" # Warn users that fix will remove user path and move it to system path
		}
		ShowPathLength $registryPathString
		$pathString = $registryPathString
	} else {
		if ($FromProcessNameOrId -ne "") {
			Write-Warning "In the context of $($process.Name) (PID $($process.Id)), PATH is different from the one stored in registry" # Warn users that there will be no fix
		} else {
			Write-Warning "In this context, PATH is different from the one stored in registry" # Warn users that fix will be only applied to current context
		}
		ShowPathLength $actualPathString
		$pathString = $actualPathString
		if ($actualPathString -like "*%*") {
			Write-Warning "Your PATH is corrupt (variables have not been properly expanded).`r`nYou may fix it by running 'GetPath -Reload'."
			#Check if registry key is expandable (with GetValueKind)
			exit -5
		} else {
			$diffMode = $true
		}
	}
	if ($pathString) {
		$pathEntries = $pathString.Split(';')
	} else {
		exit 0
	}
	$pathChecker = [System.Collections.ArrayList]@()
	$entryOrder = 1
	$driveList = PSDrive -PSProvider FileSystem | Select Name, DisplayRoot | Where {$_.DisplayRoot -ne $null}

	foreach($pathEntry in $pathEntries) {
		Write-Progress "Analyzing PATH entries" -Status "Running" -PercentComplete (($entryOrder-1)/$pathEntries.Count*100) -CurrentOperation $pathEntry
		if ($pathEntry.Contains('%')){
			$unexpandedEntry = $true
		} else {
			$unexpandedEntry = $false
		}
		$isNetworkPath = $null
		$uncPath = $null
		$pristinePath = $null
		if ($pathEntry.Length -gt 1) {
			$driveLetter = $pathEntry.SubString(0,2)
			if ($driveList -And $driveLetter -match "[a-z]{1}:") {
				$uncDrive = $driveList | where Name -eq $driveLetter.SubString(0,1) | select -ExpandProperty DisplayRoot
				if ($uncDrive -And [bool]([Uri]$uncDrive).IsUnc) {
					$isNetworkPath = $true
					$uncPath = $pathEntry.Replace($driveLetter,$uncDrive)
				}
			} elseif ($driveLetter -match "\\\\") {
				$isNetworkPath = $true
				$uncPath = $pathEntry
			}
			if ($uncPath) {
				$pristinePath = GetShortPathEntry($uncPath)
			} else {
				$pristinePath = GetShortPathEntry($pathEntry)
			}
		}
		$pathChecker.Add(@{
			EntryOrder = $entryOrder
			OriginalPath = $pathEntry
			UnexpandedEntry = $unexpandedEntry
			PristinePath = $pristinePath
			Issues = [System.Collections.ArrayList]@()
			IsNetworkPath = $isNetworkPath
			UNCPath = $uncPath
		}) | Out-Null
		$entryOrder++
	}
	Write-Progress "Analyzing PATH entries" -Status "Finished" -Completed
	DisplayPath $pathChecker $diffMode $Where
	ListDuplicates($pathChecker)
	ListIssues($pathChecker)
	
<# DEBUG Helper
	if($PSVersionTable.PSVersion.Major -gt 3) {
		$pathChecker.ForEach({[PSCustomObject]$_}) | Format-Table -AutoSize
	} else {
		$(foreach ($ht in $pathChecker){new-object PSObject -Property $ht}) | Format-Table -AutoSize	
	}
#>
}

if($PSVersionTable.PSVersion.Major -gt 2) {
	$scriptRoot = $PSScriptRoot
} else {
	$scriptRoot = split-path -parent $MyInvocation.MyCommand.Definition
}

try {
	$color = $host.ui.RawUI.ForegroundColor
	Main
} finally {
	$host.ui.RawUI.ForegroundColor = $color
}
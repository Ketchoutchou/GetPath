<#
.SYNOPSIS
GetPath helps you detect and fix issues in your PATH environment variable on Windows

.DESCRIPTION

.PARAMETER DontCheckUnexpandedDuplicates
Do not take variable-based path entries into account

.PARAMETER ProcessNameOrId
Get PATH environment variable from another running process (id or approximate name)

.PARAMETER Version
Show the current version number

.INPUTS
None. You cannot pipe objects to GetPath (yet!)

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

Param(
	[switch]$DontCheckUnexpandedDuplicates = $false,
	#[switch]$Fix = $false,
	#[switch]$FixEvenUnexpandedDuplicates = $false,
	[switch]$FromBatch = $false,
	[string]$PathExt = "",
	[string]$ProcessNameOrId = "",
	#[switch]$RestoreLongPaths = $false,
	#[switch]$ShortenAllPaths = $false,
	[switch]$TestMode = $false,
	[switch]$Verbatim = $false,
	[switch]$Version = $false,
	[string]$Where = ""
)

Set-StrictMode -Version Latest

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
 \______/  \_______|  \____/ \__|      \_______|  \____/ \__|  \__| 2.0

'@
		exit 0
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
function ShowPorcelainPath { 
	Param (
		[String]$systemPath,
		$userPath
	)

	$headerColor = "DarkGray"
	$systemPathColor = "Gray"
	$userPathColor = "White"
	
	$host.ui.RawUI.ForegroundColor = $headerColor
	echo "---------- PATH BEGIN ----------"
	if ($systemPath) {
		$host.ui.RawUI.ForegroundColor = $systemPathColor
		$systemPath.Split(';')
	}
	if ($userPath -is [String]) {
		$host.ui.RawUI.ForegroundColor = $userPathColor
		$userPath.Split(';')
	}
	$host.ui.RawUI.ForegroundColor = $headerColor
	echo "----------- PATH END -----------"
	$host.ui.RawUI.ForegroundColor = "Gray"
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
	if ($expandedPathLength -gt 4094) {
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
		if ($pathEntry -And (Test-Path $pathEntry -IsValid) -And (Test-Path $pathEntry)) {
			$FSO = New-Object -ComObject Scripting.FileSystemObject
			return $FSO.GetFolder($pathEntry).ShortPath
		} else {
			return $pathEntry.TrimEnd('\')
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
		if (!(Test-Path $pathCheckerEntry.OriginalPath -IsValid)){
			$pathCheckerEntry.Issues.Add("ShouldBeValid") | Out-Null
		} elseif (!(Test-Path (NewPrepareForTestPath($pathCheckerEntry.OriginalPath)))){
			$pathCheckerEntry.Issues.Add("MustExist") | Out-Null
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
		$flags += "N"
	} else {
		$flags += "-"
	}
	if ($pathEntry.UnexpandedEntry) {
		$flags += "%"
	} else {
		$flags += "-"
	}
	#userpath
	#duplicates
	#issues
	$flags += "`t"
	$flags
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
	if ($FromBatch) {
		$pathExtEntries = $PathExt.Split(';')
	} else {
		$pathExtEntries = $env:PathExt.Split(';')
	}
	foreach ($pathCheckerEntry in $pathChecker) {
		$colorBefore = $host.ui.RawUI.ForegroundColor
		if (!$Verbatim) {
			$prefix = GetPathPrefix $pathCheckerEntry
		} else {
			$prefix = $null
		}
		
		if ($where -ne "") {
			$filelist = @()
			$searchPattern = $pathCheckerEntry.PristinePath
			$foundFiles = gci "$searchPattern\$where.*" -Force -Name -File -Depth 0
			if ($foundFiles){
				foreach ($pathExtEntry in $pathExtEntries) {
					$fileList += gci $searchPattern -Force -Name -File -Include $where$pathExtEntry -Depth 0
				}
			}
			$fileList += gci $searchPattern -Force -Name -File -Include "$where*" -Exclude $pathExtEntries.Replace('.',"*.") -Depth 0
			if ($fileList) {
				$host.ui.RawUI.ForegroundColor = "Magenta"
			}
		}
		
		if (!$diffMode) {
			echo "$prefix$($pathCheckerEntry.OriginalPath)"
			$i = registryPathEntriesCount
		} else {
			if ($i -lt $registryPathEntriesCount -And $pathCheckerEntry.OriginalPath -eq $registryPathEntries[$i]) {
				echo "$prefix$($pathCheckerEntry.OriginalPath)"
				$i++
			} else {
				$indexInRegistry = $registryPathEntries.IndexOf($pathCheckerEntry.OriginalPath)
				if ($indexInRegistry -ne -1 -And $indexInRegistry -gt $i) {
					for ($j = $i; $j -lt $indexInRegistry; $j++) {
						$host.ui.RawUI.ForegroundColor = "Red"
						echo "`t`t($($registryPathEntries[$j])) (not present in this context, only in registry)"
					}
					$i = $indexInRegistry + 1
					$host.ui.RawUI.ForegroundColor = $colorBefore
					echo "$prefix$($pathCheckerEntry.OriginalPath) (only present in this context, not in registry)"
				}
			}
		}
		$host.ui.RawUI.ForegroundColor = "DarkGray"
		if ($where -ne "" -And $fileList) {
			foreach ($file in $fileList) {
				echo `t`t`t$file
			}
		}
		$host.ui.RawUI.ForegroundColor = $colorBefore
	}
	if ($i -lt $registryPathEntriesCount) {
		for ($j = $i; $j -lt $registryPathEntriesCount; $j++) {
			$host.ui.RawUI.ForegroundColor = "Red"
			echo "`t`t($($registryPathEntries[$j])) (not present in this context, only in registry)"
		}
		$host.ui.RawUI.ForegroundColor = $colorBefore
	}
}
function Main {
	ShowVersion
	PSVersionCheck
	$systemRegistryPathString = GetPathFromRegistry "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
	$userRegistryPathString = GetPathFromRegistry "HKCU:\Environment"
	
	$actualPathString = $env:PATH

	if ($ProcessNameOrId -ne "") {
		if($PSVersionTable.PSVersion.Major -gt 2) {
			$scriptRoot = $PSScriptRoot
		} else {
			$scriptRoot = $MyInvocation.MyCommand.Path
		}
		$getExternalProcessPathExecutable = "GetExternalProcessEnv.exe"
		if (Test-Path $scriptRoot\$getExternalProcessPathExecutable) {
			$externalProcessPathString = $null
			$exitCode = -666
			if ($ProcessNameOrId -match '^\d+$') {
				$process = Get-Process -Id $ProcessNameOrId
				Write-Warning "Analyzing process $($process.Name) (PID $($process.Id))"
				$externalProcessPathString = & $scriptRoot\$getExternalProcessPathExecutable $ProcessNameOrId
				$exitCode = $LASTEXITCODE
			}
			if ($exitCode -eq 0) {
				$actualPathString = $externalProcessPathString
			} else {
				try {
					$foundProcesses = Get-Process $ProcessNameOrId -ErrorAction 'Stop' | Select Name, Id
				} catch {
					$processList = Get-Process | Select Name, Id
					$foundProcesses = @()
					foreach ($process in $processList) {
						if ($process.Name -like "*$ProcessNameOrId*") {
							$foundProcesses += $process
						}
					}
				}
				if ($foundProcesses -is [array]) {
					if ($foundProcesses.Length -gt 1) {
						Write-Warning "Multiple processes found with name containing '$ProcessNameOrId':"
						$foundProcesses
						exit -1
					}
					if ($foundProcesses.Length -eq 0) {
						Write-Warning "No process found with name containing '$ProcessNameOrId'"
						exit -1
					}
					$process = $foundProcesses[0]
				} else {
					$process = $foundProcesses
				}
				Write-Warning "Analyzing process $($process.Name) (PID $($process.Id))"
				$externalProcessPathString = $ $scriptRoot\$getExternalProcessPathExecutable $process.Id
				$actualPathString = $externalProcessPathString
			}
		} else {
			Write-Warning "$getExternalProcessPathExecutable not found. Cannot get PATH of an external process."
			exit -1
		}
	}
	
	if ($TestMode) {
		$systemRegistryPathString = @'
		C:\windows;;s:\;::;\\;s;s:;C:\WINDOWS; C:\windows\;C:/windows/;c:\>;c:;c;c:\windows\\;c:\fdsf\\;\\\;c:\<;%USERPROFILE%\desktop;"c:\program files (x86)"\google;"c:\program files (x86)"\google2;  ;C:\Users\Ketchoutchou\Desktop;c:\doesnotexist;c:\dOesnotexist\ ; c:\program files;C:\windows*;*;?;|;c:|windows;c:\windows?;c:\program files (x86);%SystemRoot%\system32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SYSTEMROOT%\System32\WindowsPowerShell\v1.0;C:\ProgramData\Oracle\Java\javapath;C:\Program Files (x86)\NVIDIA Corporation\PhysX\Common;D:\Logiciels\Utilitaires\InPath;"c:\program files (x86)"; 
'@
		$userRegistryPathString = @'
C:\userpath
'@
		$actualPathString = [System.Environment]::ExpandEnvironmentVariables($(JoinSystemAndUserPath $systemRegistryPathString $userRegistryPathString))
	}
	
	$registryPathString = JoinSystemAndUserPath $systemRegistryPathString $userRegistryPathString

	$diffMode = $false
	if (([System.Environment]::ExpandEnvironmentVariables($registryPathString)) -eq $actualPathString) {
		if ($userRegistryPathString -eq "") {
			Write-Warning "User PATH environment variable is defined but empty" # Warn users that fix will remove user path and move it to system path
		}
		ShowPathLength $registryPathString
		#ShowPorcelainPath $systemRegistryPathString $userRegistryPathString
		$pathString = $registryPathString
	} else {
		if ($ProcessNameOrId -ne "") {
			Write-Warning "In the context of $($process.Name) (PID $($process.Id)), PATH is different from the one store in registry" # Warn users that there will be no fix
		} else {
			Write-Warning "In this context, PATH is different from the one store in registry" # Warn users that fix will be only applied to current context
		}
		ShowPathLength $actualPathString
		#ShowPorcelainPath $actualPathString
		$pathString = $actualPathString
		$diffMode = $true
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

try {
	$color = $host.ui.RawUI.ForegroundColor
	Main
} finally {
	$host.ui.RawUI.ForegroundColor = $color
}

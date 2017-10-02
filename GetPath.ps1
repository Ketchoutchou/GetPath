Param(
	[switch]$DontCheckUnexpandedDuplicates = $false,
	#[switch]$Fix = $false,
	#[switch]$FixEvenUnexpandedDuplicates = $false,
	[switch]$TestMode = $false
)

Set-StrictMode -Version Latest
$color = $host.ui.RawUI.ForegroundColor

function PSVersionCheck {
	$currentVersion = $PSVersionTable.PSVersion.Major
	if($currentVersion -lt 5){
		Write-Warning "You are using PowerShell $currentVersion.0. Consider upgrading to PowerShell 5.0 at least."
	}
}
function GetPathFromRegistry { 
	Param (
		[parameter(Mandatory=$true)] [String]$regPath
	)
	
	(Get-Item $regPath).GetValue('PATH','',[Microsoft.WIN32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
}
function JoinSystemAndUserPath {
	Param (
		[String]$systemPath,
		[String]$userPath
	)
	
	if (!$systemPath.EndsWith(';') -And $systemPath -And $userPath) {
		$separator = ";"
	} else {
		$separator = ""
	}
	$systemPath + "$separator" + $userPath
}
function ShowPorcelainPath { 
	Param (
		[String]$systemPath,
		[String]$userPath
	)

	$headerColor = "DarkGray"
	$systemPathColor = "Gray"
	$userPathColor = "Yellow"
	
	$host.ui.RawUI.ForegroundColor = $headerColor
	echo "---------- PATH BEGIN ----------"
	if ($systemPath) {
		$host.ui.RawUI.ForegroundColor = $systemPathColor
		$systemPath.Split(';')
	}
	if ($userPath) {
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
		# Should stop ?
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
			return (Get-Item $pathEntry).FullName
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
					if ($pathCheckerEntry.UnexpandedEntry) {
						$better = " - Unexpanded path entry. Use -DontCheckUnexpandedDuplicates to ignore" # if fix mode Use -FixEvenUnexpandedDuplicates to remove
					} else {
						$better = $null
					}
					echo "    ->  $(DisplayPathEntryWithOrder $pathCheckerEntry)$better"
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
			$pathCheckerEntry.Issues.Add("ShouldNotEndWithSlash") | Out-Null
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
					# Find a way to Unicode characters (ex: [char]0x21B3)
					echo "    ->  $warning"
				}
			}
		}
		$host.ui.RawUI.ForegroundColor = "Gray"
	}
}
function Main {
	PSVersionCheck
	$systemRegistryPathString = GetPathFromRegistry "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
	$userRegistryPathString = GetPathFromRegistry "HKCU:\Environment"
	
	$actualPathString = $env:PATH
	# May have to trim last semicolon on Win10

	if ($TestMode) {
		$systemRegistryPathString = @'
		C:\windows;;C:\WINDOWS; C:\windows\;C:/windows/;c:\>;c:;c;c:\windows\\;c:\fdsf\\;\\\;c:\<;%USERPROFILE%\desktop;"c:\program files (x86)"\google;"c:\program files (x86)"\google2;  ;C:\Users\Ketchoutchou\Desktop;c:\doesnotexist;c:\dOesnotexist\ ; c:\program files;C:\windows*;*;?;|;c:|windows;c:\windows?;c:\program files (x86);%SystemRoot%\system32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SYSTEMROOT%\System32\WindowsPowerShell\v1.0;C:\ProgramData\Oracle\Java\javapath;C:\Program Files (x86)\NVIDIA Corporation\PhysX\Common;D:\Logiciels\Utilitaires\InPath;"c:\program files (x86)"; 
'@
		$userRegistryPathString = @'
C:\userpath
'@
	}

	$registryPathString = JoinSystemAndUserPath $systemRegistryPathString $userRegistryPathString
	
	if ($userRegistryPathString) {
		Write-Warning "User defined PATH environment variable is not recommended" # Warn users that fix will remove user path and move it to system path
	}
	if ($TestMode -Or ([System.Environment]::ExpandEnvironmentVariables($registryPathString)) -eq $actualPathString) {
		ShowPathLength $registryPathString
		ShowPorcelainPath $systemRegistryPathString $userRegistryPathString
		$pathString = $registryPathString
	} else {
		Write-Warning "PATH has been modified in this context (different from PATH stored in registry)" # Warn users that fix will be only applied to current context
		ShowPathLength $actualPathString
		ShowPorcelainPath $actualPathString
		$pathString = $actualPathString
	}
	if ($pathString) {
		$pathEntries = $pathString.Split(';')
	} else {
		exit 0
	}
	$pathChecker = [System.Collections.ArrayList]@()
	$entryOrder = 1
	foreach($pathEntry in $pathEntries) {
		if ($pathEntry.Contains('%')){
			$unexpandedEntry = $true
		} else {
			$unexpandedEntry = $false
		}
		$pathChecker.Add(@{
			EntryOrder = $entryOrder
			OriginalPath = $pathEntry
			UnexpandedEntry = $unexpandedEntry
			PristinePath = GetShortPathEntry($pathEntry)
			Issues = [System.Collections.ArrayList]@()
		}) | Out-Null
		$entryOrder++
	}
	ListDuplicates($pathChecker)
	ListIssues($pathChecker)
	
<# DEBUG Helper
	if($PSVersionTable.PSVersion.Major -gt 2) {
		$pathChecker.ForEach({[PSCustomObject]$_}) | Format-Table -AutoSize
	} else {
		$(foreach ($ht in $pathChecker){new-object PSObject -Property $ht}) | Format-Table -AutoSize	
	}
#>
}

Main
$host.ui.RawUI.ForegroundColor = $color
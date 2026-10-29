#
# build accdb from source
#
param(
    [string]$SourceDir,
	[string]$TargetDir = "",
	[string]$FileName = "", # empty = name from vcs options
    [string]$VcsAddInPath = "" # empty = use default path (installed version)
   
)

# Check if the script is running under a Windows service account (SYSTEM, NETWORK SERVICE, LOCAL SERVICE)
$serviceAccounts = @('SYSTEM', 'NETWORK SERVICE', 'LOCAL SERVICE')
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
if ($serviceAccounts | Where-Object { $currentUser -match $_ }) {
    Write-Warning "Warning: This script is running under a Windows service account ($currentUser). Microsoft Access should not be executed as a service!"
}
else {
    Write-Host "Running script as user: $currentUser"
}

[string]$tempFileName = "VcsBuildTempApp"
[string]$accdbFileName = $tempFileName
if ($FileName -gt "") {
    $accdbFileName = $FileName
}

$curDir = $(Get-Location)
$accdbPath = "$curDir\$accdbFileName.accdb"

# Stop only the Access process started by this script (never other MSACCESS.EXE instances)
Add-Type -Namespace VcsBuild -Name Win32 -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@

function Get-AccessProcessId {
    param($AccessApp)
    [uint32]$processId = 0
    [void][VcsBuild.Win32]::GetWindowThreadProcessId([System.IntPtr]$AccessApp.hWndAccessApp(), [ref]$processId)
    return [int]$processId
}

function Stop-AccessProcess {
    param([int]$ProcessId)
    if ($ProcessId -gt 0) {
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    }
}

# open/create access file
$access = New-Object -ComObject Access.Application
$access.Visible = $true
$accessProcessId = Get-AccessProcessId -AccessApp $access
if (-not (Test-Path $accdbPath)) {    
    $access.NewCurrentDatabase($accdbPath)
} 
else {
	$access.OpenCurrentDatabase($accdbPath)
}




[string]$addInProcessPath = ""
if ($VcsAddInPath -gt "") {
    $addInProcessPath = [System.IO.Path]::ChangeExtension($VcsAddInPath, "").TrimEnd('.')   
}
else {
    $appdata = $env:APPDATA
    $addInFolder = Join-Path $appdata "MSAccessVCS"
    $addInProcessPath = Join-Path $addInFolder "Version Control"
}

$addInPattern = "$addInProcessPath.accd[ae]"

if (-not (Test-Path $addInPattern)) {
    Write-Host "msaccess-vcs add-in not found: $addInPattern"
    Write-Host "Please install msaccess-vcs add-in first."
    exit 1
}

if (
    -not ([System.IO.Path]::IsPathRooted($SourceDir)) -or
    ($SourceDir -match "^[\\\/]") # "\source" or "/source"
) {
    $SourceDir = Join-Path -Path (Get-Location) -ChildPath $SourceDir.TrimStart('\','/','.')
}

Write-Host "Add-in path: $addInProcessPath"
Write-Host "Current path: $curDir"
Write-Host "Source: $SourceDir"
Write-Host "TargetDir: $TargetDir"
Write-Host ""

Write-Host "Start msaccess-vcs build (BuildHeadless)"
$buildResult = $null
try {
    $json = $access.Run("$addInProcessPath.API", "BuildHeadless", "$SourceDir")
    $buildResult = $json | ConvertFrom-Json
}
catch {
    Write-Host "BuildHeadless call failed: $($_.Exception.Message)"
}

Start-Sleep -Seconds 1
Write-Host "Close Access " -NoNewline
try { $access.Quit(2) } catch { }
Write-Host "." -NoNewline
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
Remove-Variable access
[GC]::Collect()
Write-Host "." -NoNewline
[GC]::WaitForPendingFinalizers()

# MSACCESS.EXE may stay resident after Quit - stop the process started by this script
Start-Sleep -Seconds 1
Stop-AccessProcess -ProcessId $accessProcessId
Start-Sleep -Seconds 2

Write-Host " completed"
Write-Host ""

# copy build log to TargetDir (also on failure, as a diagnostic artifact)
if ($buildResult -and $buildResult.logPath -and (Test-Path $buildResult.logPath) -and ($TargetDir -gt "")) {
    New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
    Copy-Item -Path $buildResult.logPath -Destination $TargetDir -Force
    Write-Host "Build log copied to: $TargetDir"
}

if (-not $buildResult -or -not $buildResult.success) {
    Write-Host "Build failed"
    if (-not $buildResult) {
        Write-Host "   (no result returned by BuildHeadless)"
    }
    else {
        Write-Host "   $($buildResult.error)"
        if ($buildResult.logPath) { Write-Host "   Log: $($buildResult.logPath)" }
    }
    exit 1
}

$builtFilePath = $buildResult.databasePath
$builtFileName = [System.IO.Path]::GetFileName($builtFilePath)
Write-Host "Built: $builtFileName ($builtFilePath)"
Write-Host "Errors: $($buildResult.errorCount), Warnings: $($buildResult.warningCount), Duration: $([math]::Round($buildResult.durationMs / 1000, 1)) s"
Write-Host "Log: $($buildResult.logPath)"

# Wait for a file to become openable with an exclusive lock, or time out.
# $access.Quit() can return before the MSACCESS process has fully exited and
# released its lock on the built database, which makes Copy-Item fail
# intermittently with "being used by another process".
function Wait-FileUnlocked {
    param(
        [string]$Path,
        [int]$TimeoutSeconds = 60
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $stream.Close()
            $stream.Dispose()
            return $true
        }
        catch {
            Start-Sleep -Milliseconds 500
        }
    }
    return $false
}

# copy file to TargetDir
if ([string]::IsNullOrEmpty($FileName)) {
    $FileName = $builtFileName
}

$targetFilePath = $builtFilePath
$builtFilePathDir = [System.IO.Path]::GetDirectoryName($builtFilePath)
if (($TargetDir -gt "") -and ($TargetDir -ne  $builtFilePathDir) ) {
	Write-Host "Copy accdb to $TargetDir"
	New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null

    # Make sure the built database is no longer locked before copying it. If it is
    # still locked after the grace window, stop the Access process started by this script.
    if (-not (Wait-FileUnlocked -Path $builtFilePath -TimeoutSeconds 60)) {
        Write-Warning "Built file still locked after 60s. Stopping Access process $accessProcessId..."
        Stop-AccessProcess -ProcessId $accessProcessId
        $null = Wait-FileUnlocked -Path $builtFilePath -TimeoutSeconds 10
    }

    $copyAttempts = 0
    while ($true) {
        $copyAttempts++
        try {
            Copy-Item -Path $builtFilePath -Destination "$TargetDir\$FileName" -Force -ErrorAction Stop
            break
        }
        catch {
            if ($copyAttempts -ge 5) { throw }
            Write-Warning "Copy attempt $copyAttempts failed: $($_.Exception.Message). Retrying..."
            Start-Sleep -Seconds 2
        }
    }
	Write-Host ""
	$targetFilePath = "$TargetDir\$FileName"
} elseif ($FileName -ne $builtFileName) {
	Rename-Item -Path ".\$builtFileName" -NewName $FileName -Force
}

$tempFilePath = Join-Path -Path $curDir -ChildPath ([System.IO.Path]::ChangeExtension($tempFileName, "accdb"))
if (Test-Path $tempFilePath) {
    Remove-Item -Path $tempFilePath -Force  
}

return "$targetFilePath"

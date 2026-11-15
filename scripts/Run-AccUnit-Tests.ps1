#
# run AccUnit tests
#
param(
    [string]$AccdbPath,
    [string]$AccUnitAddInPath = "", # empty = use default path (installed version)
    [string]$JUnitXmlPath = "" # empty = don't write JUnit XML test results (requires AccUnitLoader with JUnit XML support)
)

if ([string]::IsNullOrEmpty($AccdbPath) ) {
    Write-Host "No Access database file specified."
    Write-Host "Please specify the path to the Access database file."
    exit 1
}

if (-not ([System.IO.Path]::IsPathRooted($AccdbPath)) ) {
    $AccdbPath = Join-Path -Path (Get-Location) -ChildPath $AccdbPath
    if ($AccdbPath -match '[\\/][.][\\/]')
    {
        $AccdbPath = $AccdbPath -replace '[\\/][.][\\/]', '\'
    }
}

if ([string]::IsNullOrEmpty($AccUnitAddInPath)) {
    # Default path for the AccUnit add-in
    $appdata = $env:APPDATA
    $addInFolder = Join-Path $appdata "Microsoft\AddIns"
    $AccUnitAddInPath = Join-Path $addInFolder "AccUnitLoader.accda"
}
elseif (-not ([System.IO.Path]::IsPathRooted($AccUnitAddInPath)) ) {
    $AccUnitAddInPath = Join-Path -Path (Get-Location) -ChildPath $AccUnitAddInPath
    if ($AccUnitAddInPath -match '[\\/][.][\\/]')
    {
        $AccUnitAddInPath = $AccUnitAddInPath -replace '[\\/][.][\\/]', '\'
    }
}

if (-not [string]::IsNullOrEmpty($JUnitXmlPath) -and -not ([System.IO.Path]::IsPathRooted($JUnitXmlPath)) ) {
    # resolve here, otherwise Access resolves it against its own working directory
    $JUnitXmlPath = Join-Path -Path (Get-Location) -ChildPath $JUnitXmlPath
    if ($JUnitXmlPath -match '[\\/][.][\\/]')
    {
        $JUnitXmlPath = $JUnitXmlPath -replace '[\\/][.][\\/]', '\'
    }
}

$addInProcedureCallRoot = [System.IO.Path]::ChangeExtension($AccUnitAddInPath, "").TrimEnd('.')
if (-not (Test-Path $AccUnitAddInPath)) {
    Write-Host "AccUnit add-in not found: $AccUnitAddInPath"
    Write-Host "Please install AccUnitLoader add-in first."
    exit 1
}

Write-Host "Add-in path: $AccUnitAddInPath"
Write-Host "File to test: $AccdbPath"
if (-not [string]::IsNullOrEmpty($JUnitXmlPath)) {
    Write-Host "JUnit XML file: $JUnitXmlPath"
}
Write-Host ""

$access = New-Object -ComObject Access.Application
$access.Visible = $true
$access.OpenCurrentDatabase($AccdbPath)

Write-Host "Run Tests ..." -NoNewline
if ([string]::IsNullOrEmpty($JUnitXmlPath)) {
    # keep the call without arguments: works with AccUnitLoader versions without JUnit XML support
    $result = $access.Run("$addInProcedureCallRoot.AutomatedTestRun")
}
else {
    # Application.Run only passes positional arguments => pass the first three as well
    # 3 = LogFile + DebugPrint (default of AutomatedTestRun)
    $result = $access.Run("$addInProcedureCallRoot.AutomatedTestRun", "", 3, $false, $JUnitXmlPath)
}
Write-Host " completed"
Write-Host "Tests success: $result"
Write-Host "Test result:"

$logFile = "$AccdbPath.AccUnit.log"
if (Test-Path $logFile) {
    $logContent = Get-Content $logFile | Where-Object { $_.Trim() -ne "" }
    if ($logContent.Count -ge 9) {
        $resultBlock = $logContent[-9..-1]
    } else {
        $resultBlock = $logContent
    }
    Write-Host ($resultBlock -join "`n")
} else {
    Write-Host "Log file not found: $logFile"
}
Write-Host ""

$jUnitXmlFile = $null
if (-not [string]::IsNullOrEmpty($JUnitXmlPath)) {
    if (Test-Path $JUnitXmlPath) {
        $jUnitXmlFile = $JUnitXmlPath
        Write-Host "JUnit XML test results: $jUnitXmlFile"
    }
    else {
        Write-Warning "JUnit XML file not created: $JUnitXmlPath (AccUnitLoader with JUnit XML support required, see AccessCodeLib/AccUnit#84)"
    }
    Write-Host ""
}

Start-Sleep -Seconds 1

Write-Host "Close Access " -NoNewline
$access.Quit(2)
Write-Host "." -NoNewline
Start-Sleep -Seconds 1
Write-Host "." -NoNewline
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
Remove-Variable access
[GC]::Collect()
Write-Host "." -NoNewline
[GC]::WaitForPendingFinalizers()
Write-Host " completed"
Write-Host ""

return [PSCustomObject]@{
    Success = $result
    LogFile = $logFile
    JUnitXmlFile = $jUnitXmlFile
}

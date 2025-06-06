#
# build accdb from source
#
param(
    [string]$AccdbPath,
    [string]$AccUnitAddInPath = "" # empty = use default path (installed version)
   
)

if (-not ([System.IO.Path]::IsPathRooted($AccdbPath)) ) {
    $AccdbPath = Join-Path -Path (Get-Location) -ChildPath $AccdbPath
    if ($AccdbPath -match '[\\/][.][\\/]')
    {
        $AccdbPath = $AccdbPath -replace '[\\/][.][\\/]', '\'
    }
}

if (($AccUnitAddInPath -gt "") -and -not ([System.IO.Path]::IsPathRooted($AccUnitAddInPath)) ) {
    $AccUnitAddInPath = Join-Path -Path (Get-Location) -ChildPath $AccUnitAddInPath
    if ($AccUnitAddInPath -match '[\\/][.][\\/]')
    {
        $AccUnitAddInPath = $AccUnitAddInPath -replace '[\\/][.][\\/]', '\'
    }
}


[string]$addInProcessPath = ""
if ($AccUnitAddInPath -gt "") {
    $addInProcessPath = [System.IO.Path]::ChangeExtension($AccUnitAddInPath, "").TrimEnd('.')   
}
else {
    $appdata = $env:APPDATA
    $addInFolder = Join-Path $appdata "Microsoft\AddIns"
    $addInProcessPath = Join-Path $addInFolder "AccUnitLoader"
    $AccUnitAddInPath = "$addInProcessPath.accda"
}
if (-not (Test-Path $AccUnitAddInPath)) {
    Write-Host "msaccess-vcs add-in not found: $AccUnitAddInPath"
    Write-Host "Please install AccUnitLoader add-in first."
    exit 1
}

Write-Host "Add-in path: $AccUnitAddInPath"
Write-Host "File to test: $AccdbPath"
Write-Host ""

$access = New-Object -ComObject Access.Application
$access.Visible = $true
$access.OpenCurrentDatabase($AccdbPath)

Write-Host "Run Tests ..." -NoNewline
$result = $access.Run("$addInProcessPath.AutomatedTestRun")
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
}

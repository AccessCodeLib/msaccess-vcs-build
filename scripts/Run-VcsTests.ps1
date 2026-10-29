#
# run msaccess-vcs tests (RunTestsHeadless) on a temporary copy of an accdb file
#
param(
    [string]$AccdbPath,
    [string]$Filter = "", # test filter, e.g. "-slow" (see msaccess-vcs Testing wiki)
    [string]$VcsAddInPath = "", # empty = use default path (installed version)
    [string]$TargetDir = "" # empty = don't copy the JUnit XML and the log
)

if ([string]::IsNullOrEmpty($AccdbPath) -or -not (Test-Path $AccdbPath)) {
    Write-Host "Access database file not found: $AccdbPath"
    exit 1
}

[string]$addInProcessPath = ""
if ($VcsAddInPath -gt "") {
    $addInProcessPath = [System.IO.Path]::ChangeExtension($VcsAddInPath, "").TrimEnd('.')
}
else {
    $addInProcessPath = Join-Path (Join-Path $env:APPDATA "MSAccessVCS") "Version Control"
}

if (-not (Test-Path "$addInProcessPath.accd[ae]")) {
    Write-Host "msaccess-vcs add-in not found: $addInProcessPath.accd[ae]"
    Write-Host "Please install msaccess-vcs add-in first."
    exit 1
}

# The tests run on a copy: msaccess-vcs may add modules (modTestAssert) and write its export files,
# which must not end up in the built file.
$curDir = (Get-Location).Path
$workDir = Join-Path $curDir ("VcsTests_" + (Get-Date -Format "yyyyMMddHHmmss"))
New-Item -Path $workDir -ItemType Directory -Force | Out-Null
$testAccdbPath = Join-Path $workDir ([System.IO.Path]::GetFileName($AccdbPath))
Copy-Item -Path $AccdbPath -Destination $testAccdbPath -Force

Write-Host "Add-in path: $addInProcessPath"
Write-Host "File to test: $testAccdbPath (copy of $AccdbPath)"
Write-Host "Filter: $Filter"
Write-Host ""

# Stop only the Access process started by this script (never other MSACCESS.EXE instances)
if (-not ('VcsBuild.Win32' -as [type])) {
    Add-Type -Namespace VcsBuild -Name Win32 -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@
}

$access = New-Object -ComObject Access.Application
$access.Visible = $true
[uint32]$accessProcessId = 0
[void][VcsBuild.Win32]::GetWindowThreadProcessId([System.IntPtr]$access.hWndAccessApp(), [ref]$accessProcessId)

$testResult = $null
try {
    $access.OpenCurrentDatabase($testAccdbPath)
    Write-Host "Run msaccess-vcs tests (RunTestsHeadless) ..." -NoNewline
    $json = $access.Run("$addInProcessPath.API", "RunTestsHeadless", "$Filter")
    Write-Host " completed"
    $testResult = $json | ConvertFrom-Json
}
catch {
    Write-Host ""
    Write-Host "RunTestsHeadless call failed: $($_.Exception.Message)"
}

Write-Host "Close Access " -NoNewline
try { $access.Quit(2) } catch { }
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
Remove-Variable access
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
Write-Host "." -NoNewline
Start-Sleep -Seconds 1
if ($accessProcessId -gt 0) {
    Stop-Process -Id $accessProcessId -Force -ErrorAction SilentlyContinue
}
Start-Sleep -Seconds 2
Write-Host " completed"
Write-Host ""

# copy JUnit XML and log to TargetDir (also when tests fail)
$junitTargetPath = ""
$logTargetPath = ""
if ($testResult -and ($TargetDir -gt "")) {
    New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
    if ($testResult.junitPath -and (Test-Path $testResult.junitPath)) {
        Copy-Item -Path $testResult.junitPath -Destination $TargetDir -Force
        $junitTargetPath = Join-Path $TargetDir ([System.IO.Path]::GetFileName($testResult.junitPath))
        Write-Host "JUnit XML copied to: $junitTargetPath"
    }
    if ($testResult.logPath -and (Test-Path $testResult.logPath)) {
        Copy-Item -Path $testResult.logPath -Destination $TargetDir -Force
        $logTargetPath = Join-Path $TargetDir ([System.IO.Path]::GetFileName($testResult.logPath))
        Write-Host "Test log copied to: $logTargetPath"
    }
}

# remove the temporary folder created above
if ((Test-Path $workDir) -and ($workDir.StartsWith($curDir))) {
    Remove-Item -Path $workDir -Recurse -Force -ErrorAction SilentlyContinue
}

$success = [bool]($testResult -and $testResult.allPassed -and -not $testResult.cancelled)
if ($success) {
    Write-Host "msaccess-vcs tests passed"
}
else {
    Write-Host "msaccess-vcs tests failed" -ForegroundColor Red
    if ($testResult -and $testResult.cancelled) { Write-Host "   (test run was cancelled)" }
    if ($testResult -and $testResult.error) { Write-Host "   $($testResult.error)" }
    if ($testResult -and $testResult.logPath) { Write-Host "   Log: $($testResult.logPath)" }
}
Write-Host ""

return [PSCustomObject]@{
    Success   = $success
    JunitFile = $junitTargetPath
    LogFile   = $logTargetPath
}

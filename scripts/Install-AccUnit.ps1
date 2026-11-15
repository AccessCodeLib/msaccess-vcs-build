#
# download AccUnitLoader.accda and install it (compiled) into %APPDATA%\Microsoft\AddIns
# same steps as AccUnitLoader_Install.vbs of the AccUnit release (compiled variant + add-in registration)
# + trusted location for the add-in folder
#
param(
    [string]$AccUnitUrl = "https://api.github.com/repos/AccessCodeLib/AccUnit/releases/latest",
    [string]$TargetRootDir = "", # download folder, empty = use current directory
    [bool]$SetTrustedLocation = $true # set trusted location for add-in folder (if not already trusted)
)

Write-Host "Download url: $AccUnitUrl"

$TargetOfficeApp = "Access"
. "$PSScriptRoot/Get-GitHubHeaders.ps1"
$headers = Get-GitHubHeaders -Uri $AccUnitUrl
$release = Invoke-RestMethod -Uri $AccUnitUrl -Headers $headers

# zip url
$asset = $release.assets | Where-Object { $_.name -like "AccUnitLoader*$TargetOfficeApp.zip" } | Select-Object -First 1
$zipUrl = $asset.browser_download_url

# save as
$zipFile = "AccUnitLoader.zip"

# download file
Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile

Write-Host "zip file downloaded from $zipUrl to $zipFile"

# extract to download folder
if ([string]::IsNullOrEmpty($TargetRootDir)) {
    $TargetRootDir = (Get-Location).Path
}
else {
    if (-not ([System.IO.Path]::IsPathRooted($TargetRootDir))) {
        $TargetRootDir = Join-Path -Path (Get-Location) -ChildPath $TargetRootDir
        if ($TargetRootDir -match '[\\/][.][\\/]')
        {
            $TargetRootDir = $TargetRootDir -replace '[\\/][.][\\/]', '\'
        }
    }
}
$downloadFolder = Join-Path -Path $TargetRootDir -ChildPath "AccUnit"
Expand-Archive -Path $zipFile -DestinationPath $downloadFolder -Force

$addInFileName = "AccUnitLoader.accda"
[string]$sourceAddInPath = Join-Path $downloadFolder $addInFileName

# check if add-in file exists
if (-not (Test-Path $sourceAddInPath)) {
    Write-Error "Add-in file not found: $sourceAddInPath"
    exit 1
}

# install into add-in folder: the AccUnit dll path (lib folder) is next to the add-in file
$addInFolder = Join-Path $env:APPDATA "Microsoft\AddIns"
if (-not (Test-Path $addInFolder)) {
    New-Item -ItemType Directory -Path $addInFolder | Out-Null
}
[string]$addInPath = Join-Path $addInFolder $addInFileName
[string]$accdbPath = "$addInPath.accdb" # file to compile

if ($SetTrustedLocation)
{
    # check if the add-in folder is already a trusted location (or a subfolder of one with AllowSubfolders)
    $trustedLocationsRegPath = "HKCU:\Software\Microsoft\Office\16.0\Access\Security\Trusted Locations"
    $addInFolderCompare = $addInFolder.TrimEnd('\')
    $existingTrustedLocation = $null
    if (Test-Path $trustedLocationsRegPath) {
        foreach ($location in Get-ChildItem -Path $trustedLocationsRegPath) {
            $locationPath = $location.GetValue("Path")
            if ([string]::IsNullOrEmpty($locationPath)) {
                continue
            }
            $locationPath = [System.Environment]::ExpandEnvironmentVariables($locationPath).TrimEnd('\')
            $allowSubfolders = $location.GetValue("AllowSubfolders") -eq 1
            if ($addInFolderCompare -ieq $locationPath -or
                ($allowSubfolders -and $addInFolderCompare.StartsWith($locationPath + '\', [System.StringComparison]::OrdinalIgnoreCase))) {
                $existingTrustedLocation = $location.PSChildName
                break
            }
        }
    }

    if ($existingTrustedLocation) {
        Write-Host "Add-in folder is already a trusted location ($existingTrustedLocation): $addInFolder"
    }
    else {
        Write-Host "Set trusted location: $addInFolder"
        & "$PSScriptRoot/Set-TrustedLocation.ps1" "AccUnit-add-in-folder" "$addInFolder"
    }
}

# remove old add-in files
foreach ($fileToDelete in @($addInPath, (Join-Path $addInFolder "lib\AccessCodeLib.AccUnit.tlb"), $accdbPath)) {
    if (Test-Path $fileToDelete) {
        Remove-Item -Path $fileToDelete -Force
    }
}

Copy-Item -Path $sourceAddInPath -Destination $accdbPath -Force

$access = New-Object -ComObject Access.Application
$access.Visible = $true

# run precompile procedures
$access.OpenCurrentDatabase($accdbPath)
[void]$access.Run("CheckAccUnitTypeLibFile")
[void]$access.Run("ExportAccUnitFiles")
$access.CloseCurrentDatabase()

# compile (accde)
$accessType = $access.GetType()
[void]$accessType.InvokeMember(
    "SysCmd",
    "InvokeMethod",
    $null,
    $access,
    @(603, $accdbPath, $addInPath)
)

if (Test-Path $accdbPath) {
    Remove-Item -Path $accdbPath -Force
}

if (-not (Test-Path $addInPath)) {
    $access.Quit(2)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
    Write-Error "Compiled add-in file not created: $addInPath"
    exit 1
}

# register menu add-in (USysRegInfo)
Write-Host "Register add-in"
$profileRegPath = "HKCU\SOFTWARE\Microsoft\Office\" + $access.Version + "\Access"
$wsh = New-Object -ComObject WScript.Shell
$addInDb = $access.DBEngine.OpenDatabase($addInPath)
$rst = $addInDb.OpenRecordset("select Subkey, ValName, Type, Value from USysRegInfo where ValName > '' Order By ValName", 8) # dbOpenForwardOnly
while (-not $rst.EOF) {
    $subKey = [string]$rst.Fields.Item("Subkey").Value
    $valName = [string]$rst.Fields.Item("ValName").Value
    $regType = [int]$rst.Fields.Item("Type").Value
    $itemValue = $rst.Fields.Item("Value").Value
    if ($itemValue -is [System.DBNull]) {
        $itemValue = ""
    }
    if ("$itemValue".Contains("|ACCDIR")) {
        $itemValue = $addInDb.Name
    }

    $regName = $subKey.Replace("HKEY_CURRENT_ACCESS_PROFILE", $profileRegPath)
    if ($valName -gt "") {
        $regName = $regName + "\" + $valName
    }
    switch ($regType) {
        1 { $wsh.RegWrite($regName, "$itemValue", "REG_SZ") }
        4 { $wsh.RegWrite($regName, [int]$itemValue, "REG_DWORD") }
        0 { $wsh.RegWrite($regName, "$itemValue") }
        default {
            Write-Error "Registry type not supported: $regType ($regName)"
            exit 1
        }
    }
    $rst.MoveNext()
}
$rst.Close()
$addInDb.Close()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($rst)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($addInDb)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsh)

$access.Quit(2)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
Remove-Variable access
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "AccUnitLoader installed: $addInPath"

$result = [PSCustomObject]@{
    AddInPath = "$addInPath"
    AddInFolder = "$addInFolder"
    AddInFileName = "$addInFileName"
    Success = $true
}
return $result

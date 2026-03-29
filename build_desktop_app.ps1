$ErrorActionPreference = "Stop"

$python = Join-Path $env:LocalAppData "Programs\Python\Python312\python.exe"
if (-not (Test-Path $python)) {
    throw "Python 3.12 not found at $python"
}

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$desktopDir = [Environment]::GetFolderPath("Desktop")
$distDir = Join-Path $projectRoot "dist"
$buildDir = Join-Path $projectRoot "build"
$iconPath = Join-Path $projectRoot "assets\\app_icon.ico"
$specPath = Join-Path $projectRoot "CNInfoReportCollector.spec"

$appName = "CNInfoReportCollector"
$desktopExePath = Join-Path $desktopDir "$appName.exe"

if (-not (Test-Path $iconPath)) {
    throw "Icon file not found at $iconPath"
}
if (-not (Test-Path $specPath)) {
    throw "Spec file not found at $specPath"
}

Get-ChildItem $desktopDir | Where-Object {
    $_.Name -like "CNInfo*" -or $_.Name -like "*年报采集器*"
} | ForEach-Object {
    if ($_.FullName -ne $desktopExePath) {
        Remove-Item $_.FullName -Recurse -Force
    }
}

& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --distpath $distDir `
    --workpath $buildDir `
    $specPath

$builtExePath = Join-Path $distDir "$appName.exe"
$finalExePath = $builtExePath

try {
    Copy-Item $builtExePath $desktopExePath -Force
    $finalExePath = $desktopExePath
} catch {
    Write-Warning "Desktop exe is in use. The newly built file is still available at $builtExePath"
}

Write-Host "Desktop app created:"
Write-Host $finalExePath

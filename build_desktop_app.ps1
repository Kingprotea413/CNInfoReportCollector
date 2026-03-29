$ErrorActionPreference = "Stop"

$python = Join-Path $env:LocalAppData "Programs\Python\Python312\python.exe"
if (-not (Test-Path $python)) {
    throw "Python 3.12 not found at $python"
}

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$desktopDir = [Environment]::GetFolderPath("Desktop")
$distDir = Join-Path $projectRoot "dist"
$buildDir = Join-Path $projectRoot "build"

$appName = "CNInfoReportCollector"
$desktopExePath = Join-Path $desktopDir "$appName.exe"

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
    --windowed `
    --onefile `
    --name $appName `
    --distpath $distDir `
    --workpath $buildDir `
    --specpath $projectRoot `
    (Join-Path $projectRoot "app.py")

Copy-Item (Join-Path $distDir "$appName.exe") $desktopExePath -Force

Write-Host "Desktop app created:"
Write-Host $desktopExePath

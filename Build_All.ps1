#requires -Version 5.1
<#+
.SYNOPSIS
  Build a true Windows EXE + Portable ZIP + signed-hash-manifest + Inno Setup installer.
.DESCRIPTION
  Must run on Windows with Python 3.11+ and Inno Setup 6.
  Run Install_Build_Tools.ps1 first if the build toolchain is incomplete.
  Never erases the user's installed database or documents.
#>
param([switch]$SkipDependencies)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
$Version = '10.0.5'
$BuildDir = Join-Path $Root '.build'
$EnvDir = Join-Path $Root '.build-venv'
$BinDir = Join-Path $Root 'build'
$DistDir = Join-Path $Root 'dist'
$Source = Join-Path $Root 'برنامج تسليم الرخام الاحترافي.py'
$Icon = Join-Path $Root 'app_icon.ico'
$Assets = Join-Path $Root 'assets'
$InnoScript = Join-Path $Root 'بناء_ملف_التثبيت.iss'
$ExpectedSetup = Join-Path $DistDir "Setup_MarbleDeliveryPro_v$Version.exe"
$ExpectedPortable = Join-Path $DistDir "MarbleDeliveryPro_v$Version`_Portable.zip"
$env:PYTHONUTF8 = '1'

function Assert-ExitCode([string]$Task) {
    if ($LASTEXITCODE -ne 0) { throw "$Task failed (exit $LASTEXITCODE)" }
}
function Find-InnoCompiler {
    $candidatePaths = @()
    if (${env:ProgramFiles(x86)}) { $candidatePaths += Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe' }
    if ($env:ProgramFiles) { $candidatePaths += Join-Path $env:ProgramFiles 'Inno Setup 6\ISCC.exe' }
    if ($env:LOCALAPPDATA) { $candidatePaths += Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe' }
    $onPath = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($onPath) { return $onPath.Source }
    foreach ($item in $candidatePaths) { if ($item -and (Test-Path $item)) { return $item } }
    return $null
}
function Find-Python {
    if (Get-Command py.exe -ErrorAction SilentlyContinue) { return @('py.exe','-3') }
    if (Get-Command python.exe -ErrorAction SilentlyContinue) { return @('python.exe') }
    throw 'Python not found. Install Python 3.11+ and retry.'
}

if (-not (Test-Path $Source)) { throw "Missing app source: $Source" }
if (-not (Test-Path $Icon)) { throw "Missing icon: $Icon" }
if (-not (Test-Path $Assets)) { throw "Missing assets: $Assets" }
$iscc = Find-InnoCompiler
if (-not $iscc) {
    throw 'Inno Setup 6 ISCC.exe not installed. Run .\Install_Build_Tools.ps1 then retry; installer build is required, not optional.'
}

Write-Host "[1/7] Marble Delivery Pro $Version - preparing Windows build" -ForegroundColor Cyan
$null = Find-Python
if (-not (Test-Path (Join-Path $EnvDir 'Scripts\python.exe'))) {
    if (Get-Command py.exe -ErrorAction SilentlyContinue) {
        & py.exe -3 -m venv $EnvDir
    } else {
        & python.exe -m venv $EnvDir
    }
    Assert-ExitCode 'venv creation'
}
$Py = Join-Path $EnvDir 'Scripts\python.exe'
if (-not $SkipDependencies) {
    Write-Host '[2/7] Installing Python dependencies and PyInstaller...' -ForegroundColor Cyan
    & $Py -m pip install --upgrade pip
    Assert-ExitCode 'pip upgrade'
    & $Py -m pip install -r (Join-Path $Root 'requirements-build.txt')
    Assert-ExitCode 'dependency installation'
} else {
    Write-Host '[2/7] Dependency install skipped (cached environment requested).' -ForegroundColor DarkYellow
}

Write-Host '[3/7] Running regression tests in temporary isolated data...' -ForegroundColor Cyan
& $Py -m unittest discover -s (Join-Path $Root 'tests') -v
Assert-ExitCode 'regression test suite'
Write-Host 'Running source self-test (all six report formats)...' -ForegroundColor Cyan
New-Item $BuildDir,$DistDir -ItemType Directory -Force | Out-Null
$TestDir = Join-Path $BuildDir 'self-test-data'
if (Test-Path $TestDir) { Remove-Item $TestDir -Recurse -Force }
New-Item $TestDir -ItemType Directory | Out-Null
$originalDataDir = $env:MARBLE_DELIVERY_DATA_DIR
try {
    $env:MARBLE_DELIVERY_DATA_DIR = $TestDir
    & $Py $Source --self-test
    Assert-ExitCode 'source self-test'
} finally { $env:MARBLE_DELIVERY_DATA_DIR = $originalDataDir }

Write-Host '[4/7] Building native Windows launcher...' -ForegroundColor Cyan
$exeBuild = Join-Path $BinDir 'MarbleDeliveryPro'
$PyInstallerWork = Join-Path $BuildDir 'pyinstaller'
$PyInstallerSpec = Join-Path $BuildDir 'spec'
$AddIcon = "$Icon;."
$AddAssets = "$Assets;assets"
& $Py -m PyInstaller --noconfirm --clean --noconsole --onedir --name MarbleDeliveryPro `
    --icon $Icon --distpath $BinDir --workpath $PyInstallerWork --specpath $PyInstallerSpec `
    --add-data $AddIcon --add-data $AddAssets `
    --collect-all reportlab --hidden-import arabic_reshaper --hidden-import bidi.algorithm --hidden-import xlrd $Source
Assert-ExitCode 'PyInstaller build'
$Exe = Join-Path $exeBuild 'MarbleDeliveryPro.exe'
if (-not (Test-Path $Exe)) { throw "Missing generated EXE: $Exe" }

Write-Host '[5/7] Running compiled EXE in a separate test database...' -ForegroundColor Cyan
$ExeTestDir = Join-Path $BuildDir 'exe-self-test-data'
New-Item $ExeTestDir -ItemType Directory -Force | Out-Null
try {
    $env:MARBLE_DELIVERY_DATA_DIR = $ExeTestDir
    $exeTestProc = Start-Process -FilePath $Exe -ArgumentList '--self-test' -Wait -PassThru
    if ($exeTestProc.ExitCode -ne 0) { throw 'Compiled EXE self-test failed' }
} finally { $env:MARBLE_DELIVERY_DATA_DIR = $originalDataDir }

Write-Host '[6/7] Creating the portable bundle and compiling real Setup.exe...' -ForegroundColor Cyan
if (Test-Path $ExpectedPortable) { Remove-Item $ExpectedPortable -Force }
Compress-Archive -Path (Join-Path $exeBuild '*') -DestinationPath $ExpectedPortable -CompressionLevel Optimal
if (Test-Path $ExpectedSetup) { Remove-Item $ExpectedSetup -Force }
& $iscc "/O$DistDir" $InnoScript
Assert-ExitCode 'Inno Setup compilation'
if (-not (Test-Path $ExpectedSetup)) { throw "Setup.exe not produced: $ExpectedSetup" }
if ((Get-Item $ExpectedSetup).Length -lt 1048576) { throw 'Installer suspiciously small; refusing release.' }

Write-Host '[7/7] SHA256 verification + release manifest...' -ForegroundColor Cyan
$hashFile = Join-Path $DistDir "SHA256_v$Version.txt"
@($ExpectedSetup,$ExpectedPortable) | ForEach-Object {
    $hash = (Get-FileHash $_ -Algorithm SHA256).Hash
    "$hash  $([IO.Path]::GetFileName($_))"
} | Set-Content $hashFile -Encoding UTF8
@($ExpectedSetup,$ExpectedPortable,$hashFile) | ForEach-Object {
    Get-Item $_ | Select-Object Name,Length,FullName
} | Format-Table -AutoSize
Write-Host 'PASS: Source self-test, EXE self-test, Portable ZIP, and Setup.exe created.' -ForegroundColor Green

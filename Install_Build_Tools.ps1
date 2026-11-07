#requires -Version 5.1
$ErrorActionPreference = 'Stop'
Write-Host 'Checking Windows build toolchain for Marble Delivery Pro...' -ForegroundColor Cyan
if (-not (Get-Command py.exe -ErrorAction SilentlyContinue) -and -not (Get-Command python.exe -ErrorAction SilentlyContinue)) {
    Write-Host 'Python 3.11+ is required. Installing via winget...' -ForegroundColor Yellow
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { throw 'Install Python from https://python.org then retry.' }
    winget install -e --id Python.Python.3.13 --accept-source-agreements --accept-package-agreements
    if ($LASTEXITCODE -ne 0) { throw 'Python installation failed.' }
}
$possible = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
)
$iscc = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if (-not $iscc -and -not ($possible | Where-Object { Test-Path $_ })) {
    Write-Host 'Installing Inno Setup 6 using winget...' -ForegroundColor Yellow
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw 'Inno Setup 6 required: download at https://jrsoftware.org/isinfo.php'
    }
    winget install -e --id JRSoftware.InnoSetup --accept-source-agreements --accept-package-agreements
    if ($LASTEXITCODE -ne 0) { throw 'Inno Setup installation failed. Install manually then retry.' }
}
Write-Host 'Toolchain provisioned. Now run .\Build_All.ps1' -ForegroundColor Green

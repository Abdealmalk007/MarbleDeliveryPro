#requires -Version 5.1
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
$Source = Join-Path $Root 'برنامج تسليم الرخام الاحترافي.py'
if (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 $Source
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    & python $Source
} else {
    throw 'Python غير موجود. شغّل Install_Requirements.ps1 بعد تثبيت Python.'
}
if ($LASTEXITCODE -ne 0) {
    throw "انتهى البرنامج برمز خطأ: $LASTEXITCODE"
}

#requires -Version 5.1
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

if (Get-Command py -ErrorAction SilentlyContinue) {
    $Python = @('py','-3')
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $Python = @('python')
} else {
    throw 'Python غير موجود أو غير مضاف إلى PATH.'
}

function Run-Python([string[]]$PythonArgs) {
    if ($Python[0] -eq 'py') { & py -3 @PythonArgs } else { & python @PythonArgs }
    if ($LASTEXITCODE -ne 0) { throw "فشل الأمر: $($PythonArgs -join ' ')" }
}

Write-Host 'تثبيت/إصلاح مكتبات نظام تسليم الرخام...' -ForegroundColor Cyan
Run-Python -PythonArgs @('-m','pip','install','--upgrade','pip')
Run-Python -PythonArgs @('-m','pip','install','-r',(Join-Path $Root 'requirements.txt'),'--upgrade')
Run-Python -PythonArgs @('-c','import openpyxl, reportlab, PIL, arabic_reshaper, bidi, xlrd; print("OK - جميع المكتبات جاهزة")')
Write-Host 'اكتمل بنجاح.' -ForegroundColor Green
Read-Host 'اضغط Enter للإغلاق'

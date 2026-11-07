#requires -Version 5.1
<#
Publish the reviewed clean-source repo and all three release assets.
No network changes happen unless the user explicitly types PUBLISH.
Requires GitHub CLI (gh), a connected GitHub account, and built release assets.
#>
param([string]$RepositoryName = 'MarbleDeliveryPro')
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$release = 'E:\ProTech\Releases\MarbleDeliveryPro_v10.0.5_PUBLIC'
$files = @(
  (Join-Path $release 'Setup_MarbleDeliveryPro_v10.0.5.exe'),
  (Join-Path $release 'MarbleDeliveryPro_v10.0.5_Portable.zip'),
  (Join-Path $release 'MarbleDeliveryPro_v10.0.5_Source.zip')
)
foreach ($file in $files) { if (-not (Test-Path -LiteralPath $file)) { throw "Missing release asset: $file" } }
if (-not (Get-Command gh.exe -ErrorAction SilentlyContinue)) { throw 'Install GitHub CLI: https://cli.github.com/ and run gh auth login.' }
& gh auth status
if ($LASTEXITCODE -ne 0) { throw 'Run gh auth login first.' }
$login = (& gh api user --jq .login).Trim()
if ($LASTEXITCODE -ne 0 -or -not $login) { throw 'Could not read GitHub account.' }
$fullName = "$login/$RepositoryName"
Write-Host "Account: $login; public repository: $fullName"
$confirmation = Read-Host 'Type PUBLISH to publish source and release assets publicly'
if ($confirmation -cne 'PUBLISH') { Write-Host 'Cancelled. No GitHub changes.'; exit 0 }
& gh repo view $fullName *> $null
if ($LASTEXITCODE -ne 0) {
    & gh repo create $fullName --public --source $root --remote origin
    if ($LASTEXITCODE -ne 0) { throw 'Unable to create repository.' }
}
if (-not ((& git -C $root remote) -contains 'origin')) {
    & git -C $root remote add origin "https://github.com/$fullName.git"
    if ($LASTEXITCODE -ne 0) { throw 'Failed to set Git origin.' }
}
& git -C $root push -u origin main
if ($LASTEXITCODE -ne 0) { throw 'Git push failed.' }
& gh release view v10.0.5 --repo $fullName *> $null
if ($LASTEXITCODE -eq 0) {
    & gh release upload v10.0.5 @files --repo $fullName --clobber
} else {
    & gh release create v10.0.5 @files --repo $fullName --title 'MarbleDeliveryPro 10.0.5' --notes-file (Join-Path $root 'RELEASE_NOTES_AR.md')
}
if ($LASTEXITCODE -ne 0) { throw 'GitHub release failed.' }
Write-Host "SUCCESS https://github.com/$fullName/releases/tag/v10.0.5"

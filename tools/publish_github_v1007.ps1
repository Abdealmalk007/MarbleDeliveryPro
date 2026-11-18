#requires -Version 5.1
$ErrorActionPreference='Stop'
$root='E:\ProTech\MarbleDeliveryPro_v10.0.7'
$repo='E:\ProTech\MarbleDeliveryPro_GitHub_v10.0.7'
$release='E:\ProTech\Releases\MarbleDeliveryPro_v10.0.7_PUBLIC'
$py=Join-Path $root '.build-venv\Scripts\python.exe'
& $py (Join-Path $root 'tools\stage_github_public_v1007.py')
if ($LASTEXITCODE -ne 0) { throw 'Source staging failed' }
& $py (Join-Path $root 'tools\package_local_release_v1007.py')
if ($LASTEXITCODE -ne 0) { throw 'Release packaging failed' }
Copy-Item -LiteralPath (Join-Path $root 'RELEASE_NOTES_AR.md') -Destination (Join-Path $release 'RELEASE_NOTES_v10.0.7_AR.md') -Force
Copy-Item -LiteralPath (Join-Path $root 'docs\QA_REPORT_V1007_AR.md') -Destination (Join-Path $release 'QA_REPORT_v10.0.7.md') -Force
Set-Location $repo
& git fetch origin main
if ($LASTEXITCODE -ne 0) { throw 'Unable to refresh app repo' }
& git status --short
& git add -A
if ($LASTEXITCODE -ne 0) { throw 'git add failed' }
$stagedFiles=@(& git diff --cached --name-only)
if ($stagedFiles.Count -lt 3) { throw 'Too few staged changes; refusing to publish' }
foreach ($filename in $stagedFiles) {
  if ($filename -match '(?i)\.(db|sqlite|sqlite3|exe|zip|pdf|xlsx|csv|log|bak)$|^\.env') {
    throw "Private or binary file staged: $filename"
  }
}
& git commit -m 'release: MarbleDeliveryPro v10.0.7 wider accurate PDF and complete analytics'
if ($LASTEXITCODE -ne 0) { throw 'git commit failed' }
& git push origin HEAD:main
if ($LASTEXITCODE -ne 0) { throw 'git push failed; changes local only' }
Write-Output 'APP_SOURCE_PUSHED'
$existing=(& gh release view v10.0.7 -R Abdealmalk007/MarbleDeliveryPro --json tagName -q .tagName 2>$null)
if ($existing -eq 'v10.0.7') { throw 'Release v10.0.7 already exists; refuse accidental overwrite' }
$assets=@(
  (Join-Path $release 'Setup_MarbleDeliveryPro_v10.0.7.exe'),
  (Join-Path $release 'MarbleDeliveryPro_v10.0.7_Portable.zip'),
  (Join-Path $release 'MarbleDeliveryPro_v10.0.7_Source.zip'),
  (Join-Path $release 'SHA256SUMS.txt')
)
foreach ($asset in $assets) { if(-not (Test-Path -LiteralPath $asset)) {throw "Missing $asset"} }
& gh release create v10.0.7 @assets -R Abdealmalk007/MarbleDeliveryPro --title 'MarbleDeliveryPro v10.0.7 - Reports & Analytics Preview' --prerelease --notes-file (Join-Path $root 'RELEASE_NOTES_AR.md') --target main
if ($LASTEXITCODE -ne 0) { throw 'Release upload failed' }
& gh release view v10.0.7 -R Abdealmalk007/MarbleDeliveryPro --json url,assets,publishedAt --jq '{url:.url,publishedAt:.publishedAt,assets:[.assets[].name]}'
if ($LASTEXITCODE -ne 0) { throw 'Cannot verify published release' }
Write-Output 'GITHUB_PUBLIC_RELEASE_PASS'

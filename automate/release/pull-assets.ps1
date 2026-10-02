#!/usr/bin/env pwsh
# pull-assets.ps1 — Windows twin of pull-assets.sh. Downloads + restores the working
# binaries from the 'working-binaries' GitHub Release into src/bin, .mock, ribbon templates.
$ErrorActionPreference = "Stop"
Set-Location (& git rev-parse --show-toplevel)

$Repo       = if ($env:OBT_REPO) { $env:OBT_REPO } else { "epicentre-msf/outbreak-tools" }
$Tag        = "working-binaries"
$BundleName = "working-binaries.tar.gz"
$TmpDir     = Join-Path ([System.IO.Path]::GetTempPath()) ("obt-" + [System.Guid]::NewGuid().ToString())

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw "gh CLI not found." }
New-Item -ItemType Directory -Path $TmpDir | Out-Null
try {
  Write-Host "==> downloading $BundleName from release '$Tag'"
  gh release download $Tag -R $Repo -p $BundleName -D $TmpDir --clobber
  Write-Host "==> extracting (overwrites src/bin, .mock, ribbon templates)"
  tar -xzf (Join-Path $TmpDir $BundleName)
  Write-Host "Done."
} finally {
  Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue
}

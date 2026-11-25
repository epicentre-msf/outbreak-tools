#!/usr/bin/env pwsh
# pull-assets.ps1 [-MockRibbons] — Windows twin of pull-assets.sh. Downloads + restores
# the working binaries from the 'working-binaries' GitHub Release into src/bin, .mock,
# the ribbon templates and trads/.
#
# -MockRibbons restores .mock/ and the ribbon templates alone: src/bin, the trads
# workbooks and the test inputs on this disk are left as they are.
param([switch]$MockRibbons)
$ErrorActionPreference = "Stop"
Set-Location (& git rev-parse --show-toplevel)

$Repo       = if ($env:OBT_REPO) { $env:OBT_REPO } else { "epicentre-msf/outbreak-tools" }
$Tag        = "working-binaries"
$BundleName = "working-binaries.tar.gz"
$TmpDir     = Join-Path ([System.IO.Path]::GetTempPath()) ("obt-" + [System.Guid]::NewGuid().ToString())
# What -MockRibbons restores. Keep it equal to the list in push-assets.ps1.
$MockRibbonPaths = @(".mock",
                     "ribbons/_ribbontemplate_main.xlsm", "ribbons/_ribbontemplate_dev.xlsm",
                     "ribbons/_ribbontemplate_main.xlsb", "ribbons/_ribbontemplate_dev.xlsb")

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw "gh CLI not found." }
New-Item -ItemType Directory -Path $TmpDir | Out-Null
try {
  Write-Host "==> downloading $BundleName from release '$Tag'"
  gh release download $Tag -R $Repo -p $BundleName -D $TmpDir --clobber
  $Bundle = Join-Path $TmpDir $BundleName
  if (-not $MockRibbons) {
    Write-Host "==> extracting (overwrites src/bin, .mock, ribbon templates, trads workbooks)"
    tar -xzf $Bundle
    Write-Host "Done."
    return
  }

  # tar stops on a name the bundle does not hold, so each path is looked up in the
  # listing first and only the ones found are asked for.
  $listing = @(tar -tzf $Bundle)
  $wanted  = @()
  foreach ($p in $MockRibbonPaths) {
    if ($listing | Where-Object { $_ -eq $p -or $_.StartsWith("$p/") } | Select-Object -First 1) { $wanted += $p }
  }
  if ($wanted.Count -eq 0) { throw "the bundle holds no mock and no ribbon template." }

  Write-Host "==> extracting $($wanted -join ', ')"
  Write-Host "    (src/bin, the trads workbooks and the test inputs are left as they are)"
  tar -xzf $Bundle @wanted
  if ($LASTEXITCODE -ne 0) { throw "tar could not extract the mocks and the ribbon templates." }
  Write-Host "Done. Mocks and ribbon templates restored."
} finally {
  Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue
}

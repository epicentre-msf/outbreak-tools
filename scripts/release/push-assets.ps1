#!/usr/bin/env pwsh
# push-assets.ps1 — Windows twin of push-assets.sh. Bundles the working binaries and the
# trads translation workbooks, and uploads them to the pinned 'working-binaries' GitHub
# Release. Creates it on first run.
$ErrorActionPreference = "Stop"
Set-Location (& git rev-parse --show-toplevel)

$Repo   = if ($env:OBT_REPO) { $env:OBT_REPO } else { "epicentre-msf/outbreak-tools" }
$Tag    = "working-binaries"
$Bundle = Join-Path ([System.IO.Path]::GetTempPath()) "working-binaries.tar.gz"
$Paths  = @("src/bin", ".mock", "ribbons/_ribbontemplate_main.xlsb", "ribbons/_ribbontemplate_dev.xlsb",
            "trads/designer_translations*.xlsx")

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw "gh CLI not found." }

# PowerShell hands tar the literal argument, so a wildcard entry is expanded here.
$existing = @()
foreach ($p in $Paths) {
  if ($p -match '\*') {
    $hits = @(Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue |
              ForEach-Object { (Resolve-Path -Relative $_.FullName) -replace '^\.[\\/]', '' -replace '\\', '/' })
    if ($hits.Count -gt 0) { $existing += $hits } else { Write-Host "WARN: missing $p (skipped)" }
  }
  elseif (Test-Path $p) { $existing += $p }
  else { Write-Host "WARN: missing $p (skipped)" }
}
if ($existing.Count -eq 0) { throw "no binary paths found to bundle." }

Write-Host "==> bundling: $($existing -join ', ')"
tar -czf $Bundle @existing
try {
  gh release view $Tag -R $Repo *> $null
  if ($LASTEXITCODE -ne 0) {
    Write-Host "==> creating asset-store release '$Tag'"
    gh release create $Tag -R $Repo --prerelease --title "Working binaries (mutable)" --notes "Off-git store of the current working binaries. Mutable; do not link directly. Synced via push/pull-assets."
  }
  Write-Host "==> uploading"
  gh release upload $Tag -R $Repo $Bundle --clobber
  Write-Host "Done."
} finally {
  Remove-Item $Bundle -Force -ErrorAction SilentlyContinue
}

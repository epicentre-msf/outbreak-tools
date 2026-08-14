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
    # Parked on the initial commit and left untitled on purpose — see RELEASING.md §5.
    gh release create $Tag -R $Repo --prerelease --target e2bb1f46bd01389823f65fcb632d4e9d12ce2cae --title "" --notes "Off-git store of the current working binaries, synced with scripts/release/push-assets.sh and pull-assets.sh.`n`nInfrastructure, mutable, not a download."
  }
  Write-Host "==> uploading"
  gh release upload $Tag -R $Repo $Bundle --clobber
  Write-Host "Done."

  # New binaries in the store mean the single dev pre-release is stale, so rebuild it.
  # Set OBT_SKIP_DEV_REFRESH=1 to publish binaries without republishing the dev build.
  if ($env:OBT_SKIP_DEV_REFRESH -eq "1") {
    Write-Host "==> skipping the dev-latest refresh (OBT_SKIP_DEV_REFRESH=1)"
  } else {
    Write-Host "==> refreshing the dev-latest pre-release"
    gh workflow run dev-latest.yml -R $Repo --ref dev
    if ($LASTEXITCODE -ne 0) { Write-Host "WARN: could not trigger dev-latest.yml — run it from the Actions tab." }
  }
} finally {
  Remove-Item $Bundle -Force -ErrorAction SilentlyContinue
}

#!/usr/bin/env pwsh
# push-assets.ps1 [-MockRibbons] — Windows twin of push-assets.sh. Bundles the working
# binaries and the trads translation workbooks, and uploads them to the pinned
# 'working-binaries' GitHub Release. Creates it on first run.
#
# -MockRibbons sends .mock/ and the ribbon templates alone. The store is one tarball,
# so the bundle already in the store is downloaded, those paths are replaced by the
# ones on this disk, and the bundle goes back up: src/bin, the trads workbooks and the
# test inputs stay exactly as the store had them. The preflight leaves src/bin out,
# and the dev-latest pre-release is not rebuilt.
param([switch]$MockRibbons)
$ErrorActionPreference = "Stop"
Set-Location (& git rev-parse --show-toplevel)

$Repo       = if ($env:OBT_REPO) { $env:OBT_REPO } else { "epicentre-msf/outbreak-tools" }
$Tag        = "working-binaries"
$BundleName = "working-binaries.tar.gz"
$Work       = Join-Path ([System.IO.Path]::GetTempPath()) ("obt-" + [System.Guid]::NewGuid().ToString())
$Bundle     = Join-Path $Work $BundleName
# What -MockRibbons sends: the mocks, and the ribbon templates in both formats while
# the .xlsb ones are still around.
$MockRibbonPaths = @(".mock",
                     "ribbons/_ribbontemplate_main.xlsm", "ribbons/_ribbontemplate_dev.xlsm",
                     "ribbons/_ribbontemplate_main.xlsb", "ribbons/_ribbontemplate_dev.xlsb")
# Every path the bundle holds: those, the built binaries, the translation workbooks
# and the frozen test inputs.
$Paths = @("src/bin") + $MockRibbonPaths + @("trads/designer_translations*.xlsx", "src/tests/.input")

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw "gh CLI not found." }

# The entries of $Wanted that exist under the current directory. PowerShell hands tar
# the literal argument, so a wildcard entry is expanded here.
function Get-PresentPaths([string[]]$Wanted, [switch]$Quiet) {
  $found = @()
  foreach ($p in $Wanted) {
    if ($p -match '\*') {
      $hits = @(Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue |
                ForEach-Object { (Resolve-Path -Relative $_.FullName) -replace '^\.[\\/]', '' -replace '\\', '/' })
      if ($hits.Count -gt 0) { $found += $hits } elseif (-not $Quiet) { Write-Host "WARN: missing $p (skipped)" }
    }
    elseif (Test-Path $p) { $found += $p }
    elseif (-not $Quiet) { Write-Host "WARN: missing $p (skipped)" }
  }
  $found
}

# Nothing leaves this machine until the binaries and the sources agree: no 2010
# ribbon envelope, ribbon plumbing intact, the ribbons under ribbons/ identical to
# the ones inside the workbooks, and no hand-imported class trailing its .cls.
# preflight.sh rewrites nothing unless every one of those passes. RELEASING.md §2.
# It stays one bash script rather than two: the ribbon tools it drives are bash.
if (-not (Get-Command bash -ErrorAction SilentlyContinue)) {
  throw "bash not found — the release preflight (scripts/release/preflight.sh) needs it."
}
if ($MockRibbons) { bash scripts/release/preflight.sh --no-bin } else { bash scripts/release/preflight.sh }
if ($LASTEXITCODE -ne 0) { throw "preflight failed — nothing was pushed." }
Write-Host ""

$send     = if ($MockRibbons) { $MockRibbonPaths } else { $Paths }
$existing = @(Get-PresentPaths $send)
if ($existing.Count -eq 0) { throw "no binary paths found to bundle." }

New-Item -ItemType Directory -Path $Work | Out-Null
try {
  if (-not $MockRibbons) {
    Write-Host "==> bundling: $($existing -join ', ')"
    tar -czf $Bundle @existing
    if ($LASTEXITCODE -ne 0) { throw "tar could not write the bundle." }
  } else {
    # The bundle of the store is unpacked, the paths found on this disk replace their
    # copies in it, and it is packed again. A path this disk does not have is left as
    # the store had it, so this push never takes one out of the store.
    $Store = Join-Path $Work "store"
    $Stage = Join-Path $Work "stage"
    New-Item -ItemType Directory -Path $Store, $Stage | Out-Null
    Write-Host "==> downloading $BundleName from release '$Tag' (the rest of it is kept)"
    gh release download $Tag -R $Repo -p $BundleName -D $Store --clobber
    if ($LASTEXITCODE -ne 0) {
      throw "could not download $BundleName from release '$Tag'. This push keeps what the store already holds; run a full push first if the store is empty."
    }
    tar -xzf (Join-Path $Store $BundleName) -C $Stage
    if ($LASTEXITCODE -ne 0) { throw "tar could not unpack the bundle of the store." }
    foreach ($p in $existing) {
      $old = Join-Path $Stage $p
      if (Test-Path $old) { Remove-Item $old -Recurse -Force }
    }
    # Through tar, like the bash twin, so paths and dates arrive unchanged.
    $Carry = Join-Path $Work "carry.tar"
    tar -cf $Carry @existing
    if ($LASTEXITCODE -ne 0) { throw "tar could not read the paths to send." }
    tar -xf $Carry -C $Stage
    if ($LASTEXITCODE -ne 0) { throw "tar could not place the paths to send." }

    Push-Location $Stage
    try {
      $staged = @(Get-PresentPaths $Paths -Quiet)
      Write-Host "==> replaced from this disk: $($existing -join ', ')"
      Write-Host "==> bundling: $($staged -join ', ')"
      tar -czf $Bundle @staged
      if ($LASTEXITCODE -ne 0) { throw "tar could not write the bundle." }
    } finally { Pop-Location }
  }

  gh release view $Tag -R $Repo *> $null
  if ($LASTEXITCODE -ne 0) {
    Write-Host "==> creating asset-store release '$Tag'"
    # Parked on the initial commit and left untitled on purpose — see RELEASING.md §5.
    gh release create $Tag -R $Repo --prerelease --target e2bb1f46bd01389823f65fcb632d4e9d12ce2cae --title "" --notes "Off-git store of the current working binaries, synced with scripts/release/push-assets.sh and pull-assets.sh.`n`nInfrastructure, mutable, not a download."
  }
  Write-Host "==> uploading"
  gh release upload $Tag -R $Repo $Bundle --clobber
  if ($LASTEXITCODE -ne 0) { throw "the upload failed — the store is unchanged." }

  # The binaries in the store did not move, so there is no new dev build to make.
  # dev-latest does ship a ribbon template: its next refresh, whatever starts it,
  # takes the one pushed here.
  if ($MockRibbons) {
    Write-Host "Done. Mocks and ribbon templates pushed; src/bin in the store is unchanged."
    Write-Host "==> dev-latest is left as it is"
    return
  }
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
  Remove-Item $Work -Recurse -Force -ErrorAction SilentlyContinue
}

#!/usr/bin/env bash
#
# push-assets.sh — Bundle the working binaries and upload them to the pinned
# 'working-binaries' GitHub Release (the off-git asset store). Creates the release on
# first run. Run after editing binaries in Excel to publish them off-git.
#
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

REPO="${OBT_REPO:-epicentre-msf/outbreak-tools}"
TAG="working-binaries"
BUNDLE_NAME="working-binaries.tar.gz"
# Binary paths to store (whole dirs + the two ribbon templates).
PATHS=( src/bin .mock ribbons/_ribbontemplate_main.xlsb ribbons/_ribbontemplate_dev.xlsb )

command -v gh >/dev/null 2>&1 || { echo "ERROR: gh CLI not found (brew install gh)." >&2; exit 1; }

TMPD="$(mktemp -d)"; trap 'rm -rf "$TMPD"' EXIT
BUNDLE="$TMPD/$BUNDLE_NAME"

existing=()
for p in "${PATHS[@]}"; do
  if [ -e "$p" ]; then existing+=( "$p" ); else echo "WARN: missing $p (skipped)"; fi
done
[ ${#existing[@]} -gt 0 ] || { echo "ERROR: no binary paths found to bundle." >&2; exit 1; }

echo "==> bundling: ${existing[*]}"
tar -czf "$BUNDLE" "${existing[@]}"

if ! gh release view "$TAG" -R "$REPO" >/dev/null 2>&1; then
  echo "==> creating asset-store release '$TAG'"
  gh release create "$TAG" -R "$REPO" --prerelease \
    --title "Working binaries (mutable)" \
    --notes "Off-git store of the current working binaries (src/bin + .mock + ribbon templates). Mutable; do not link directly. Synced via automate/release/{push,pull}-assets."
fi

echo "==> uploading $BUNDLE_NAME"
gh release upload "$TAG" -R "$REPO" "$BUNDLE" --clobber
echo "Done. Working binaries pushed to release '$TAG'."

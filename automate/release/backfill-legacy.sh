#!/usr/bin/env bash
#
# backfill-legacy.sh — ONE-TIME migration. Upload every historical release archive in
# releases/old/ to a single 'legacy-archive' GitHub Release, so the releases/ folder can
# later be removed from git (Phase 5). Idempotent: re-runs just re-clobber the assets.
#
# Pass DRY_RUN=yes to list what would be uploaded without touching GitHub.
#
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

REPO="${OBT_REPO:-epicentre-msf/outbreak-tools}"
TAG="legacy-archive"
SRC_DIR="releases/old"
DRY_RUN="${DRY_RUN:-no}"

[ -d "$SRC_DIR" ] || { echo "ERROR: $SRC_DIR not found." >&2; exit 1; }

# Collect the archives (zips, plus any loose master-setup xlsb that lived in old/).
shopt -s nullglob
files=( "$SRC_DIR"/*.zip "$SRC_DIR"/*.xlsb )
shopt -u nullglob
[ ${#files[@]} -gt 0 ] || { echo "ERROR: no archives found in $SRC_DIR." >&2; exit 1; }

echo "==> ${#files[@]} legacy archive(s) from $SRC_DIR:"
printf '    %s\n' "${files[@]}"

if [ "$DRY_RUN" = "yes" ]; then
  echo "DRY_RUN=yes — nothing uploaded."
  exit 0
fi

command -v gh >/dev/null 2>&1 || { echo "ERROR: gh CLI not found (brew install gh)." >&2; exit 1; }

if ! gh release view "$TAG" -R "$REPO" >/dev/null 2>&1; then
  echo "==> creating release '$TAG'"
  gh release create "$TAG" -R "$REPO" --prerelease \
    --title "Legacy release archive" \
    --notes "Archive of all OBT releases produced before the GitHub-Releases workflow. Each asset is a dated OBT-{branch}-{date}.zip (designer + setup + ribbon template)."
fi

echo "==> uploading ${#files[@]} asset(s) (clobber)"
gh release upload "$TAG" -R "$REPO" "${files[@]}" --clobber
echo "Done. ${#files[@]} archive(s) on release '$TAG'."

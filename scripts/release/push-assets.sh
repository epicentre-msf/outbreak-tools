#!/usr/bin/env bash
#
# push-assets.sh — Bundle the working binaries and upload them to the pinned
# 'working-binaries' GitHub Release (the off-git asset store). Creates the release on
# first run. Run after editing binaries in Excel to publish them off-git, and
# after editing the translation workbooks in trads/.
#
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

REPO="${OBT_REPO:-epicentre-msf/outbreak-tools}"
TAG="working-binaries"
BUNDLE_NAME="working-binaries.tar.gz"
# Binary paths to store (whole dirs, the two ribbon templates, the translation
# workbooks). A glob that matches nothing stays literal and is skipped below.
PATHS=( src/bin .mock ribbons/_ribbontemplate_main.xlsb ribbons/_ribbontemplate_dev.xlsb
        trads/designer_translations*.xlsx )

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
  # Parked on the initial commit and left untitled on purpose: GitHub orders the
  # releases page by the target commit's date and shows the tag when a release has no
  # name, which keeps this infra release at the bottom and out of the way of versions.
  # See RELEASING.md §5.
  gh release create "$TAG" -R "$REPO" --prerelease \
    --target e2bb1f46bd01389823f65fcb632d4e9d12ce2cae \
    --title "" \
    --notes "Off-git store of the current working binaries, synced with scripts/release/push-assets.sh and pull-assets.sh.

Infrastructure, mutable, not a download — it holds src/bin/, .mock/, the two ribbon templates, and the trads/ translation workbooks, bundled as a single tarball."
fi

echo "==> uploading $BUNDLE_NAME"
gh release upload "$TAG" -R "$REPO" "$BUNDLE" --clobber
echo "Done. Working binaries pushed to release '$TAG'."

# New binaries in the store mean the single dev pre-release is stale, so rebuild it.
# Set OBT_SKIP_DEV_REFRESH=1 to publish binaries without republishing the dev build.
if [ "${OBT_SKIP_DEV_REFRESH:-0}" = "1" ]; then
  echo "==> skipping the dev-latest refresh (OBT_SKIP_DEV_REFRESH=1)"
else
  echo "==> refreshing the dev-latest pre-release"
  gh workflow run dev-latest.yml -R "$REPO" --ref dev \
    || echo "WARN: could not trigger dev-latest.yml — run it from the Actions tab."
fi

#!/usr/bin/env bash
#
# pull-assets.sh — Download + restore the working binaries from the 'working-binaries'
# GitHub Release into src/bin, .mock, and the ribbon templates. Run on a fresh checkout
# or to sync the latest binaries.
#
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

REPO="${OBT_REPO:-epicentre-msf/outbreak-tools}"
TAG="working-binaries"
BUNDLE_NAME="working-binaries.tar.gz"

command -v gh >/dev/null 2>&1 || { echo "ERROR: gh CLI not found (brew install gh)." >&2; exit 1; }

TMPD="$(mktemp -d)"; trap 'rm -rf "$TMPD"' EXIT

echo "==> downloading $BUNDLE_NAME from release '$TAG'"
gh release download "$TAG" -R "$REPO" -p "$BUNDLE_NAME" -D "$TMPD" --clobber

echo "==> extracting (overwrites src/bin, .mock, ribbon templates)"
tar -xzf "$TMPD/$BUNDLE_NAME"
echo "Done. Working binaries restored."

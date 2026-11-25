#!/usr/bin/env bash
#
# pull-headless.sh [--check] [--force]
#
# Restore the headless working folder, .obt/draft/headless, from the private
# store. That folder is untracked and holds the setups, the geobases, the empty
# setup template and the drivers, so a second machine has nothing to build from
# until it has been pulled once.
#
#   --check   say whether the store holds something this disk does not, and
#             download nothing but the stamp. Prints one of:
#                 headless store: IN SYNC
#                 headless store: AHEAD    (the store was pushed from elsewhere)
#                 headless store: EMPTY    (nothing has been pushed yet)
#   --force   download and extract even when the stamp says IN SYNC
#
# What a pull overwrites: every input under .obt/draft/headless. What it never
# touches: out/ and logs/, which the push leaves out of the bundle because they
# are this machine's own run output. Push before pulling on the other machine;
# the store is one bundle and the last push wins.
#
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

CHECK=0
FORCE=0
for arg in "$@"; do
  case "$arg" in
    --check)   CHECK=1 ;;
    --force)   FORCE=1 ;;
    -h|--help) sed -n '2,22p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *)         echo "unknown option: $arg" >&2; exit 2 ;;
  esac
done

HEADLESS_DIR=".obt/draft/headless"
LOCAL_STAMP="$HEADLESS_DIR/.store-stamp"
TAG="headless-folder"
BUNDLE_NAME="headless-folder.tar.gz"
STAMP_NAME="headless-folder.stamp"

command -v gh >/dev/null 2>&1 || { echo "ERROR: gh CLI not found (brew install gh)." >&2; exit 1; }

# The store is a PRIVATE repository of the person running this, because the
# folder carries real field setups and geobases and outbreak-tools itself is
# public. The default is derived from the gh login rather than written down, so
# this script names nobody and each developer gets their own store.
REPO="${OBT_HEADLESS_REPO:-}"
if [ -z "$REPO" ]; then
  login="$(gh api user --jq .login 2>/dev/null || true)"
  [ -n "$login" ] || { echo "ERROR: gh is not logged in, and OBT_HEADLESS_REPO is not set." >&2; exit 1; }
  REPO="$login/outbreak-tools-headless"
fi

sha256_of() {
  if command -v sha256sum >/dev/null 2>&1; then sha256sum "$1" | cut -d' ' -f1
  else shasum -a 256 "$1" | cut -d' ' -f1; fi
}
stamp_sha() { sed -n 's/^bundle_sha256=//p' "$1" 2>/dev/null | head -1; }

TMPD="$(mktemp -d)"; trap 'rm -rf "$TMPD"' EXIT

# The stamp is a few hundred bytes and carries the bundle's digest, so the
# question "is there anything to pull" is answered without moving 35 MB.
STORE_STAMP=""
if gh release download "$TAG" -R "$REPO" -p "$STAMP_NAME" -D "$TMPD" --clobber >/dev/null 2>&1; then
  STORE_STAMP="$TMPD/$STAMP_NAME"
fi

if [ -z "$STORE_STAMP" ]; then
  echo "headless store: EMPTY"
  echo "    nothing has been pushed to $REPO yet."
  echo "    Push from the machine that has the folder: bash scripts/headless/push-headless.sh"
  exit 0
fi

store_sha="$(stamp_sha "$STORE_STAMP")"
local_sha="$(stamp_sha "$LOCAL_STAMP")"

if [ -n "$local_sha" ] && [ "$store_sha" = "$local_sha" ]; then
  echo "headless store: IN SYNC"
  sed 's/^/    /' "$STORE_STAMP"
  [ $CHECK -eq 1 ] && exit 0
  if [ $FORCE -eq 0 ]; then
    echo "    nothing to pull. --force downloads and extracts anyway."
    exit 0
  fi
else
  echo "headless store: AHEAD"
  sed 's/^/    /' "$STORE_STAMP"
  [ $CHECK -eq 1 ] && exit 0
fi

echo "==> downloading $BUNDLE_NAME from $REPO release '$TAG'"
gh release download "$TAG" -R "$REPO" -p "$BUNDLE_NAME" -D "$TMPD" --clobber

echo "==> extracting into $HEADLESS_DIR (out/ and logs/ are not in the bundle)"
tar -xzf "$TMPD/$BUNDLE_NAME"

# The stamp is written only after the extraction succeeded, so an interrupted
# pull leaves the disk looking out of date rather than looking current.
mkdir -p "$HEADLESS_DIR"
cp "$STORE_STAMP" "$LOCAL_STAMP"

echo "Done. The headless folder is what the store holds:"
sed 's/^/    /' "$LOCAL_STAMP"

#!/usr/bin/env bash
#
# push-headless.sh [--force] [--dry-run]
#
# Publish the headless working folder, .obt/draft/headless, to the private store
# so the other machine can pull it. The folder is untracked and holds the setups,
# the geobases, the empty setup template and the drivers.
#
#   --force     push even when the store was pushed from somewhere else since
#               this disk last synced. The store is one bundle, so this REPLACES
#               what the other machine put there.
#   --dry-run   bundle and report, upload nothing
#
# What goes up: everything under .obt/draft/headless except out/ and logs/.
# Those two are this machine's own run output, several tens of megabytes of
# regenerable workbooks and one appended verdict file per machine, so sending
# them would mean each pull clobbering the other machine's record of its runs.
#
# WHY THE STORE IS A PRIVATE REPOSITORY
# -----------------------------------------------------------------------------
# epicentre-msf/outbreak-tools is public, and this folder carries real field
# setups and two geobases. The working-binaries pattern would make every one of
# them a world-readable download. So the store is a private repository of the
# person pushing, created on the first push, and this script refuses to upload
# to a public one.
#
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

FORCE=0
DRY=0
for arg in "$@"; do
  case "$arg" in
    --force)   FORCE=1 ;;
    --dry-run) DRY=1 ;;
    -h|--help) sed -n '2,27p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *)         echo "unknown option: $arg" >&2; exit 2 ;;
  esac
done

HEADLESS_DIR=".obt/draft/headless"
LOCAL_STAMP="$HEADLESS_DIR/.store-stamp"
TAG="headless-folder"
BUNDLE_NAME="headless-folder.tar.gz"
STAMP_NAME="headless-folder.stamp"

[ -d "$HEADLESS_DIR" ] || { echo "ERROR: no $HEADLESS_DIR on this disk, nothing to push." >&2; exit 1; }
command -v gh >/dev/null 2>&1 || { echo "ERROR: gh CLI not found (brew install gh)." >&2; exit 1; }

# Derived from the gh login rather than written down, so this script names
# nobody and each developer gets their own store. OBT_HEADLESS_REPO overrides.
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
BUNDLE="$TMPD/$BUNDLE_NAME"
STAMP="$TMPD/$STAMP_NAME"

# --- the store repository ----------------------------------------------------
if ! gh repo view "$REPO" >/dev/null 2>&1; then
  if [ $DRY -eq 1 ]; then
    echo "==> $REPO does not exist; a real push would create it private"
  else
    echo "==> creating the private store repository $REPO"
    # --add-readme is load-bearing: a release needs a commit to hang on, and
    # GitHub answers "HTTP 422 ... Repository is empty" for a release in a
    # repository that has none.
    gh repo create "$REPO" --private --add-readme \
      --description "Off-git store of the outbreak-tools headless working folder (setups, geobases, drivers). Private: carries real field setups."
  fi
fi

if gh repo view "$REPO" >/dev/null 2>&1; then
  visibility="$(gh repo view "$REPO" --json visibility --jq .visibility 2>/dev/null || echo UNKNOWN)"
  if [ "$visibility" != "PRIVATE" ] && [ "${OBT_HEADLESS_ALLOW_PUBLIC:-0}" != "1" ]; then
    echo "ERROR: $REPO is $visibility." >&2
    echo "       The bundle carries real field setups and geobases, so this script only" >&2
    echo "       uploads to a private repository. Make it private, point OBT_HEADLESS_REPO" >&2
    echo "       at a private one, or set OBT_HEADLESS_ALLOW_PUBLIC=1 to override." >&2
    exit 1
  fi
fi

# --- has somebody else pushed since this disk last synced? -------------------
if gh release download "$TAG" -R "$REPO" -p "$STAMP_NAME" -D "$TMPD/store" --clobber >/dev/null 2>&1; then
  store_sha="$(stamp_sha "$TMPD/store/$STAMP_NAME")"
  local_sha="$(stamp_sha "$LOCAL_STAMP")"
  if [ -n "$store_sha" ] && [ "$store_sha" != "$local_sha" ]; then
    echo "The store holds a bundle this disk never pulled:"
    sed 's/^/    /' "$TMPD/store/$STAMP_NAME"
    if [ $FORCE -eq 0 ]; then
      echo "ERROR: pushing would replace it. Pull first, or push with --force." >&2
      echo "       bash scripts/headless/pull-headless.sh" >&2
      exit 1
    fi
    echo "WARN: --force is set, so this push replaces that bundle." >&2
  fi
fi

# --- the bundle --------------------------------------------------------------
# Member paths are repo-relative, so a pull extracts at the repository root and
# the folder lands where it was. The excludes are member paths for the same
# reason; ._* and .DS_Store keep a macOS push from carrying AppleDouble
# sidecars, and the stamp is regenerated rather than shipped.
echo "==> bundling $HEADLESS_DIR (out/ and logs/ left out)"
tar -czf "$BUNDLE" \
  --exclude="$HEADLESS_DIR/out" \
  --exclude="$HEADLESS_DIR/logs" \
  --exclude="$HEADLESS_DIR/.store-stamp" \
  --exclude='._*' \
  --exclude='.DS_Store' \
  "$HEADLESS_DIR"

members="$(tar -tzf "$BUNDLE" | grep -cv '/$' || true)"
bytes="$(wc -c < "$BUNDLE" | tr -d ' ')"
sha="$(sha256_of "$BUNDLE")"

{
  echo "pushed=$(date -u '+%Y-%m-%dT%H:%M:%SZ')"
  echo "os=$(uname -s)"
  echo "files=$members"
  echo "bundle_bytes=$bytes"
  echo "bundle_sha256=$sha"
} > "$STAMP"

echo "==> bundle: $members files, $bytes bytes"
sed 's/^/    /' "$STAMP"

if [ $DRY -eq 1 ]; then
  echo "Dry run. Nothing was uploaded."
  exit 0
fi

# --- the release -------------------------------------------------------------
# A repository created by hand, or by an older run of this script, can have no
# commit at all, and a release cannot be created against one. Seed it rather
# than failing after the bundle is already built.
if ! gh api "repos/$REPO/commits" --jq '.[0].sha' >/dev/null 2>&1; then
  echo "==> seeding $REPO with a README, so a release has a commit to hang on"
  readme="# outbreak-tools headless store

Private, mutable infrastructure. Not a download, and not a code repository: one
release, \`$TAG\`, holding one tarball of the untracked headless working folder
\`.obt/draft/headless\` from the outbreak-tools checkout.

Synced with \`scripts/headless/push-headless.sh\` and \`pull-headless.sh\`.
\`out/\` and \`logs/\` are left out on purpose: they are each machine's own run
output.

**This repository stays private.** The bundle carries real field setups and
geobases, and outbreak-tools itself is public.
"
  gh api --method PUT "repos/$REPO/contents/README.md" \
    -f message="Seed the store so a release can be created" \
    -f content="$(printf '%s' "$readme" | base64 | tr -d '\n')" >/dev/null
fi

# One pinned pre-release holding one bundle. Mutable infrastructure, not a
# download, the same shape as the working-binaries store.
if ! gh release view "$TAG" -R "$REPO" >/dev/null 2>&1; then
  echo "==> creating the store release '$TAG'"
  gh release create "$TAG" -R "$REPO" --prerelease --title "" \
    --notes "Off-git store of .obt/draft/headless, synced with scripts/headless/push-headless.sh and pull-headless.sh.

Infrastructure, mutable, not a download. One tarball holding the setups, the geobases, the empty setup template and the drivers. out/ and logs/ are left out on purpose: they are each machine's own run output."
fi

echo "==> uploading $BUNDLE_NAME and $STAMP_NAME"
gh release upload "$TAG" -R "$REPO" "$BUNDLE" "$STAMP" --clobber

cp "$STAMP" "$LOCAL_STAMP"
echo "Done. $REPO release '$TAG' holds this folder. Pull it on the other machine:"
echo "    bash scripts/headless/pull-headless.sh"

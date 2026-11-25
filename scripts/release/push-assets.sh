#!/usr/bin/env bash
#
# push-assets.sh [--mock-ribbons] [--force]
#
# Bundle the working binaries and upload them to the pinned 'working-binaries'
# GitHub Release (the off-git asset store). Creates the release on first run. Run
# after editing binaries in Excel to publish them off-git, and after editing the
# translation workbooks in trads/.
#
# --mock-ribbons sends .mock/ and the ribbon templates alone. The store is one
# tarball, so the bundle already in the store is downloaded, those paths are
# replaced by the ones on this disk, and the bundle goes back up: src/bin, the
# trads workbooks and the test inputs stay exactly as the store had them. The
# preflight leaves src/bin out, and the dev-latest pre-release is not rebuilt.
#
# --force pushes even when the preflight fails. The preflight still runs and
# prints what is red. When it fails, the ribbon sources are not rewritten and
# the dev-latest pre-release is left as it is. A forced push whose preflight
# passes is a normal push.
#
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

MOCK_RIBBONS=0
FORCE=0
for arg in "$@"; do
  case "$arg" in
    --mock-ribbons) MOCK_RIBBONS=1 ;;
    --force)        FORCE=1 ;;
    -h|--help)      sed -n '2,19p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *)              echo "unknown option: $arg" >&2; exit 2 ;;
  esac
done

REPO="${OBT_REPO:-epicentre-msf/outbreak-tools}"
TAG="working-binaries"
BUNDLE_NAME="working-binaries.tar.gz"
# What --mock-ribbons sends: the mocks, and the ribbon templates in both formats
# while the .xlsb ones are still around.
MOCK_RIBBON_PATHS=( .mock
                    ribbons/_ribbontemplate_main.xlsb ribbons/_ribbontemplate_dev.xlsb
                    ribbons/_ribbontemplate_main.xlsm ribbons/_ribbontemplate_dev.xlsm )
# Every path the bundle holds: those, the built binaries, the translation
# workbooks and the frozen test inputs. The glob is quoted so that it expands
# where it is used, on this disk or inside the bundle of the store.
PATHS=( src/bin "${MOCK_RIBBON_PATHS[@]}" 'trads/designer_translations*.xlsx' src/tests/.input )

command -v gh >/dev/null 2>&1 || { echo "ERROR: gh CLI not found (brew install gh)." >&2; exit 1; }

# present_under <dir> <path>...
# The paths that exist under <dir>, one per line. One that is missing is named
# on stderr and skipped; a glob that matches nothing counts as missing.
present_under() {
  local dir="$1" pat p
  shift
  (
    cd "$dir"
    for pat in "$@"; do
      for p in $pat; do   # unquoted on purpose: this is where a glob expands
        if [ -e "$p" ]; then printf '%s\n' "$p"; else echo "WARN: missing $p (skipped)" >&2; fi
      done
    done
  )
}

if [ $MOCK_RIBBONS -eq 1 ]; then
  SEND=( "${MOCK_RIBBON_PATHS[@]}" )
  PREFLIGHT_SCOPE="--no-bin"
else
  SEND=( "${PATHS[@]}" )
  PREFLIGHT_SCOPE=""
fi

# Nothing leaves this machine until the binaries and the sources agree: no 2010
# ribbon envelope, ribbon plumbing intact, the ribbons under ribbons/ identical to
# the ones inside the workbooks, and no hand-imported class trailing its .cls.
# preflight.sh rewrites nothing unless every one of those passes. RELEASING.md §2.
# --force is the one way past a red preflight: the checks have run and printed,
# and the push goes on with the ribbon sources left as they were.
PREFLIGHT_RED=0
if ! bash scripts/release/preflight.sh $PREFLIGHT_SCOPE; then
  if [ $FORCE -eq 0 ]; then
    echo "ERROR: the preflight failed, so nothing was pushed. --force pushes anyway." >&2
    exit 1
  fi
  PREFLIGHT_RED=1
  echo "WARN: the preflight failed and --force is set, so the push goes on." >&2
fi
echo

TMPD="$(mktemp -d)"; trap 'rm -rf "$TMPD"' EXIT
BUNDLE="$TMPD/$BUNDLE_NAME"

existing=()
while IFS= read -r p; do existing+=( "$p" ); done < <(present_under . "${SEND[@]}")
[ ${#existing[@]} -gt 0 ] || { echo "ERROR: no binary paths found to bundle." >&2; exit 1; }

if [ $MOCK_RIBBONS -eq 0 ]; then
  echo "==> bundling: ${existing[*]}"
  tar -czf "$BUNDLE" "${existing[@]}"
else
  # The bundle of the store is unpacked, the paths found on this disk replace
  # their copies in it, and it is packed again. A path this disk does not have
  # is left as the store had it, so this push never takes one out of the store.
  echo "==> downloading $BUNDLE_NAME from release '$TAG' (the rest of it is kept)"
  mkdir "$TMPD/store" "$TMPD/stage"
  if ! gh release download "$TAG" -R "$REPO" -p "$BUNDLE_NAME" -D "$TMPD/store" --clobber; then
    echo "ERROR: could not download $BUNDLE_NAME from release '$TAG'." >&2
    echo "       This push keeps what the store already holds. Run a full push first" >&2
    echo "       if the store is empty: bash scripts/release/push-assets.sh" >&2
    exit 1
  fi
  tar -xzf "$TMPD/store/$BUNDLE_NAME" -C "$TMPD/stage"
  for p in "${existing[@]}"; do rm -rf "$TMPD/stage/$p"; done
  tar -cf - "${existing[@]}" | tar -xf - -C "$TMPD/stage"

  staged=()
  while IFS= read -r p; do staged+=( "$p" ); done < <(present_under "$TMPD/stage" "${PATHS[@]}" 2>/dev/null)
  echo "==> replaced from this disk: ${existing[*]}"
  echo "==> bundling: ${staged[*]}"
  tar -czf "$BUNDLE" -C "$TMPD/stage" "${staged[@]}"
fi

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

Infrastructure, mutable, not a download — it holds src/bin/, .mock/, the ribbon templates, and the trads/ translation workbooks, bundled as a single tarball."
fi

echo "==> uploading $BUNDLE_NAME"
gh release upload "$TAG" -R "$REPO" "$BUNDLE" --clobber

if [ $MOCK_RIBBONS -eq 1 ]; then
  echo "Done. Mocks and ribbon templates pushed to release '$TAG'; src/bin there is unchanged."
else
  echo "Done. Working binaries pushed to release '$TAG'."
fi
if [ $PREFLIGHT_RED -eq 1 ]; then
  echo "WARN: this push went past a failed preflight (--force). Read its report above." >&2
fi

if [ $MOCK_RIBBONS -eq 1 ]; then
  # The binaries in the store did not move, so there is no new dev build to make.
  # dev-latest does ship a ribbon template: its next refresh, whatever starts it,
  # takes the one pushed here.
  echo "==> dev-latest is left as it is"
elif [ $PREFLIGHT_RED -eq 1 ]; then
  # The checks refused these binaries. They are stored, and the public dev build
  # is not made from them until somebody asks for it.
  echo "==> dev-latest is left as it is, because the preflight failed"
  echo "    (once the binaries are right, run dev-latest.yml from the Actions tab)"
elif [ "${OBT_SKIP_DEV_REFRESH:-0}" = "1" ]; then
  # Set OBT_SKIP_DEV_REFRESH=1 to publish binaries without republishing the dev build.
  echo "==> skipping the dev-latest refresh (OBT_SKIP_DEV_REFRESH=1)"
else
  # New binaries in the store mean the single dev pre-release is stale, so rebuild it.
  echo "==> refreshing the dev-latest pre-release"
  gh workflow run dev-latest.yml -R "$REPO" --ref dev \
    || echo "WARN: could not trigger dev-latest.yml — run it from the Actions tab."
fi

#!/usr/bin/env bash
#
# pull-assets.sh [--mock-ribbons]
#
# Download + restore the working binaries from the 'working-binaries' GitHub
# Release into src/bin, .mock, the ribbon templates and trads/. Run on a fresh
# checkout or to sync the latest binaries.
#
# --mock-ribbons restores .mock/ and the ribbon templates alone: src/bin, the
# trads workbooks and the test inputs on this disk are left as they are.
#
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

MOCK_RIBBONS=0
case "${1:-}" in
  --mock-ribbons) MOCK_RIBBONS=1 ;;
  -h|--help)      sed -n '2,10p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
  "")             ;;
  *)              echo "unknown option: $1" >&2; exit 2 ;;
esac

REPO="${OBT_REPO:-epicentre-msf/outbreak-tools}"
TAG="working-binaries"
BUNDLE_NAME="working-binaries.tar.gz"
# What --mock-ribbons restores. Keep it equal to the list in push-assets.sh.
MOCK_RIBBON_PATHS=( .mock
                    ribbons/_ribbontemplate_main.xlsb ribbons/_ribbontemplate_dev.xlsb
                    ribbons/_ribbontemplate_main.xlsm ribbons/_ribbontemplate_dev.xlsm )

command -v gh >/dev/null 2>&1 || { echo "ERROR: gh CLI not found (brew install gh)." >&2; exit 1; }

TMPD="$(mktemp -d)"; trap 'rm -rf "$TMPD"' EXIT

echo "==> downloading $BUNDLE_NAME from release '$TAG'"
gh release download "$TAG" -R "$REPO" -p "$BUNDLE_NAME" -D "$TMPD" --clobber

if [ $MOCK_RIBBONS -eq 0 ]; then
  echo "==> extracting (overwrites src/bin, .mock, ribbon templates, trads workbooks)"
  tar -xzf "$TMPD/$BUNDLE_NAME"
  echo "Done. Working binaries restored."
  exit 0
fi

# tar stops on a name the bundle does not hold, so each path is looked up in the
# listing first and only the ones found are asked for.
LISTING="$(tar -tzf "$TMPD/$BUNDLE_NAME")"
in_bundle() {
  local member
  while IFS= read -r member; do
    case "$member" in "$1"|"$1"/*) return 0 ;; esac
  done <<< "$LISTING"
  return 1
}

wanted=()
for p in "${MOCK_RIBBON_PATHS[@]}"; do
  if in_bundle "$p"; then wanted+=( "$p" ); fi
done
[ ${#wanted[@]} -gt 0 ] || { echo "ERROR: the bundle holds no mock and no ribbon template." >&2; exit 1; }

echo "==> extracting ${wanted[*]}"
echo "    (src/bin, the trads workbooks and the test inputs are left as they are)"
tar -xzf "$TMPD/$BUNDLE_NAME" "${wanted[@]}"
echo "Done. Mocks and ribbon templates restored."

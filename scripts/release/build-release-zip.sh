#!/usr/bin/env bash
#
# build-release-zip.sh <main|dev> <version>
#
# Assemble OBT-<branch>-<version>.zip (designer + setup + ribbon template) from the
# working-tree binaries, matching the legacy create_obt_zip.R payload and naming.
# The caller must ensure the binaries are present first (locally they already are;
# in CI run scripts/release/pull-assets.sh before this).
#
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

BRANCH="${1:-}"
VERSION="${2:-}"
if [ -z "$BRANCH" ] || [ -z "$VERSION" ]; then
  echo "usage: build-release-zip.sh <main|dev> <version>   (e.g. build-release-zip.sh main 2026.06.14)" >&2
  exit 2
fi

OUT_DIR="${OUT_DIR:-$(pwd)}"

# hot-fixes is released through the dev stream (matches legacy create_obt_zip.R).
[ "$BRANCH" = "hot-fixes" ] && BRANCH="dev"

# Branch -> source binaries. Legacy create_obt_zip.R packaged the non-suffixed
# designer.xlsb/setup.xlsb and relied on them being GIT-branch-versioned. The asset
# store is a single branch-agnostic snapshot, so here the _dev-suffixed files ARE the
# dev build and the non-suffixed files ARE the main/promoted build: a dev release ships
# designer_dev.xlsb/setup_dev.xlsb, a main release ships designer.xlsb/setup.xlsb.
case "$BRANCH" in
  main)
    DESIGNER="src/bin/designer/designer.xlsb"
    SETUP="src/bin/setup/setup.xlsb"
    RIBBON="ribbons/_ribbontemplate_main.xlsb"
    ;;
  dev)
    DESIGNER="src/bin/designer/designer_dev.xlsb"
    SETUP="src/bin/setup/setup_dev.xlsb"
    RIBBON="ribbons/_ribbontemplate_dev.xlsb"
    ;;
  *)
    echo "ERROR: branch must be 'main', 'dev', or 'hot-fixes' (got '$BRANCH')" >&2
    exit 2
    ;;
esac

for f in "$DESIGNER" "$SETUP" "$RIBBON"; do
  [ -f "$f" ] || { echo "ERROR: missing $f — run scripts/release/pull-assets.sh first." >&2; exit 1; }
done

# Ribbon envelope gate. WPS Office rewrites a workbook's custom UI into the Office
# 2007 envelope while copying the XML body verbatim; a 2009/07 namespace then sits
# inside a 2006-typed part and Excel drops the ribbon without a word. Shipping in the
# 2007 envelope makes that rewrite a no-op, so a release must never carry customUI14.
# Run scripts/devtools/ribbon-envelope.sh on any workbook this rejects.
for f in "$DESIGNER" "$SETUP" "$RIBBON"; do
  parts=$(unzip -Z1 "$f" 2>/dev/null || true)
  if printf '%s\n' "$parts" | grep -qx 'customUI/customUI14.xml'; then
    echo "ERROR: $f ships the Office 2010 ribbon envelope (customUI14.xml)." >&2
    echo "       Run: scripts/devtools/ribbon-envelope.sh $f" >&2
    exit 1
  fi
  ns=$(unzip -p "$f" customUI/customUI.xml 2>/dev/null | grep -oE 'office/[0-9/]+/customui' | head -1)
  if [ "$ns" != "office/2006/01/customui" ]; then
    echo "ERROR: $f has customUI.xml carrying namespace '${ns:-none}'." >&2
    echo "       A 2006-typed part must declare office/2006/01/customui, or Excel discards it." >&2
    exit 1
  fi
done

command -v zip >/dev/null 2>&1 || { echo "ERROR: 'zip' not found." >&2; exit 1; }

TMPD="$(mktemp -d)"; trap 'rm -rf "$TMPD"' EXIT
cp "$DESIGNER" "$TMPD/designer_${BRANCH}-${VERSION}.xlsb"
cp "$SETUP"    "$TMPD/setup_${BRANCH}-${VERSION}.xlsb"
cp "$RIBBON"   "$TMPD/_ribbontemplate_${BRANCH}-${VERSION}.xlsb"

ZIP="$OUT_DIR/OBT-${BRANCH}-${VERSION}.zip"
rm -f "$ZIP"
# -j: flat archive (no directory paths), matching create_obt_zip.R's flags="-j".
( cd "$TMPD" && zip -j -q "$ZIP" \
    "designer_${BRANCH}-${VERSION}.xlsb" \
    "setup_${BRANCH}-${VERSION}.xlsb" \
    "_ribbontemplate_${BRANCH}-${VERSION}.xlsb" )

echo "Built: $ZIP"
# Best-effort listing — never let a missing 'unzip' fail the build after the zip exists.
if command -v unzip >/dev/null 2>&1; then
  unzip -l "$ZIP" | sed 's/^/  /'
else
  zip -sf "$ZIP" | sed 's/^/  /'
fi

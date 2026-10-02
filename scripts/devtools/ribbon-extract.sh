#!/usr/bin/env bash
#
# ribbon-extract.sh <workbook> [--out-dir <dir>] [--overwrite]
#
# Unpacks a workbook's ribbon into an editable folder, so a button can be added
# with a text editor instead of the OfficeRibbonX editor on a Windows VM.
#
#   scripts/devtools/ribbon-extract.sh .mock/setup_mock.xlsb
#     -> ribbons/setup_mock/ribbon.xml
#        ribbons/setup_mock/images/*.png
#        ribbons/setup_mock/.source
#
# Image relationships are not written out. Every workbook in this repo names an
# image relationship after its file (Id="hide" -> images/hide.png), so
# ribbon-pack.sh rebuilds them from the folder. Adding an icon is therefore:
# drop newicon.png into images/, write image="newicon" in ribbon.xml, pack.
#
# Round-trips with scripts/devtools/ribbon-pack.sh.
set -euo pipefail

WB=""; OUT_DIR="ribbons"; OVERWRITE=0
while [ $# -gt 0 ]; do
  case "$1" in
    --out-dir)   OUT_DIR="${2:?--out-dir needs a path}"; shift 2 ;;
    --overwrite) OVERWRITE=1; shift ;;
    -h|--help)   sed -n '2,20p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    -*)          echo "unknown option: $1" >&2; exit 2 ;;
    *)           WB="$1"; shift ;;
  esac
done
[ -n "$WB" ] || { echo "usage: ribbon-extract.sh <workbook> [--out-dir <dir>] [--overwrite]" >&2; exit 2; }
[ -f "$WB" ] || { echo "ERROR: no such workbook: $WB" >&2; exit 1; }

LIST=$(unzip -Z1 "$WB" 2>/dev/null || true)

# Accept either envelope on the way in; ribbon-pack.sh is what enforces
# conventions.md §11 on the way out.
PART=""
for cand in customUI/customUI.xml customUI/customUI14.xml; do
  printf '%s\n' "$LIST" | grep -qx "$cand" && { PART="$cand"; break; }
done
[ -n "$PART" ] || { echo "ERROR: $WB carries no ribbon (no customUI part)." >&2; exit 1; }

STEM=$(basename "$WB"); STEM="${STEM%.*}"
DEST="$OUT_DIR/$STEM"
if [ -e "$DEST" ] && [ $OVERWRITE -eq 0 ]; then
  echo "ERROR: $DEST already exists. Pass --overwrite to replace it." >&2
  echo "       (it may hold ribbon edits that are not in the workbook yet)" >&2
  exit 1
fi
rm -rf "$DEST"; mkdir -p "$DEST/images"

unzip -p "$WB" "$PART" > "$DEST/ribbon.xml"

COUNT=0
while IFS= read -r img; do
  [ -z "$img" ] && continue
  unzip -p "$WB" "$img" > "$DEST/images/$(basename "$img")"
  COUNT=$((COUNT + 1))
done <<< "$(printf '%s\n' "$LIST" | grep '^customUI/images/' || true)"
[ $COUNT -gt 0 ] || rmdir "$DEST/images"

printf 'workbook=%s\npart=%s\n' "$WB" "$PART" > "$DEST/.source"

NS=$(grep -oE 'office/[0-9/]+/customui' "$DEST/ribbon.xml" | head -1)
echo "extracted $WB -> $DEST"
echo "  ribbon.xml   ($NS)"
echo "  images/      ($COUNT png)"
if [ "$NS" != "office/2006/01/customui" ]; then
  echo "WARNING: that namespace is the Office 2010 one. conventions.md §11 requires" >&2
  echo "         office/2006/01/customui, and ribbon-pack.sh will reject it. Either run" >&2
  echo "         ribbon-envelope.sh on the source workbook, or edit the xmlns in" >&2
  echo "         $DEST/ribbon.xml before packing." >&2
fi

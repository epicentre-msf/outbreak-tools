#!/usr/bin/env bash
#
# ribbon-pack.sh <ribbon-dir> --out <workbook> [--into <base-workbook>] [--overwrite]
#
# Packs a folder produced by ribbon-extract.sh back into a workbook.
#
#   scripts/devtools/ribbon-pack.sh ribbons/setup_mock --out .mock/setup_mock.xlsb --overwrite
#   scripts/devtools/ribbon-pack.sh ribbons/setup_mock --out /tmp/try.xlsb
#
# --into names the workbook supplying every non-ribbon part; it defaults to the
# workbook recorded in .source at extract time. --out is where the result goes,
# and an existing --out is only replaced when --overwrite is passed.
#
# The ribbon is written in the Office 2007 envelope (conventions.md §11):
# customUI/customUI.xml, an office/2006 extensibility relationship, and a body
# declaring office/2006/01/customui. Image relationships are rebuilt from the
# files in images/, each Id taking its filename stem.
#
# Refuses to build a workbook Excel would silently reject:
#   - ribbon.xml not well-formed
#   - a namespace other than office/2006/01/customui
#   - an image="X" with no images/X.png   <-- Excel drops the WHOLE ribbon for this
set -euo pipefail

NS06="http://schemas.microsoft.com/office/2006/01/customui"
REL06="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"

DIR=""; OUT=""; INTO=""; OVERWRITE=0
while [ $# -gt 0 ]; do
  case "$1" in
    --out)       OUT="${2:?--out needs a path}"; shift 2 ;;
    --into)      INTO="${2:?--into needs a path}"; shift 2 ;;
    --overwrite) OVERWRITE=1; shift ;;
    -h|--help)   sed -n '2,24p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    -*)          echo "unknown option: $1" >&2; exit 2 ;;
    *)           DIR="$1"; shift ;;
  esac
done
[ -n "$DIR" ] && [ -n "$OUT" ] || {
  echo "usage: ribbon-pack.sh <ribbon-dir> --out <workbook> [--into <base>] [--overwrite]" >&2; exit 2; }
[ -d "$DIR" ] || { echo "ERROR: no such folder: $DIR" >&2; exit 1; }
[ -f "$DIR/ribbon.xml" ] || { echo "ERROR: $DIR/ribbon.xml is missing." >&2; exit 1; }

# base workbook: --into, else whatever ribbon-extract.sh recorded
if [ -z "$INTO" ]; then
  [ -f "$DIR/.source" ] || { echo "ERROR: no .source in $DIR — pass --into <base-workbook>." >&2; exit 1; }
  INTO=$(grep '^workbook=' "$DIR/.source" | cut -d= -f2-)
fi
[ -f "$INTO" ] || { echo "ERROR: base workbook not found: $INTO" >&2; exit 1; }

if [ -e "$OUT" ] && [ $OVERWRITE -eq 0 ]; then
  echo "ERROR: $OUT already exists. Pass --overwrite to replace it." >&2; exit 1
fi

# ---- validate before touching anything -------------------------------------
if command -v xmllint >/dev/null 2>&1; then
  xmllint --noout "$DIR/ribbon.xml" 2>&1 || { echo "ERROR: $DIR/ribbon.xml is not well-formed XML." >&2; exit 1; }
fi

NS=$(grep -oE 'office/[0-9/]+/customui' "$DIR/ribbon.xml" | head -1 || true)
if [ "$NS" != "office/2006/01/customui" ]; then
  echo "ERROR: $DIR/ribbon.xml declares '${NS:-no customui namespace}'." >&2
  echo "       conventions.md §11 requires office/2006/01/customui, because WPS Office" >&2
  echo "       rewrites any other envelope into a form Excel discards in silence." >&2
  exit 1
fi

missing=""; used=""
for id in $(grep -oE 'image="[^"]+"' "$DIR/ribbon.xml" | sed 's/image="//;s/"//' | sort -u); do
  used="$used $id"
  [ -f "$DIR/images/$id.png" ] || missing="$missing $id"
done
if [ -n "$missing" ]; then
  echo "ERROR: ribbon.xml references images with no file in $DIR/images:" >&2
  for m in $missing; do echo "         image=\"$m\"  (expected $DIR/images/$m.png)" >&2; done
  echo "       Excel drops the entire ribbon over this, with no error shown." >&2
  exit 1
fi

if [ -d "$DIR/images" ]; then
  for p in "$DIR"/images/*.png; do
    [ -e "$p" ] || continue
    stem=$(basename "$p" .png)
    printf '%s' " $used " | grep -q " $stem " || echo "note: images/$stem.png is not referenced by ribbon.xml (it will still be packed)"
  done
fi

# ---- build ------------------------------------------------------------------
STAGE=$(mktemp -d)
OUTDIR=$(cd "$(dirname "$OUT")" && pwd); ABS_OUT="$OUTDIR/$(basename "$OUT")"

# The package is built beside --out and moved into place at the very end.
# --out and --into name the same workbook whenever a folder is packed back over
# the workbook it was extracted from, and cp refuses to copy a file onto itself;
# building aside also leaves --out untouched when a later step fails.
BUILD="$ABS_OUT.packing.$$"
trap 'rm -rf "$STAGE"; rm -f "$BUILD"' EXIT
mkdir -p "$STAGE/customUI/_rels" "$STAGE/customUI/images" "$STAGE/_rels"

cp "$INTO" "$BUILD"

cp "$DIR/ribbon.xml" "$STAGE/customUI/customUI.xml"

# image relationships, rebuilt from the folder: Id = filename stem
{
  printf '%s' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  printf '\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  if [ -d "$DIR/images" ]; then
    for p in "$DIR"/images/*.png; do
      [ -e "$p" ] || continue
      stem=$(basename "$p" .png)
      cp "$p" "$STAGE/customUI/images/$stem.png"
      printf '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/%s.png"/>' "$stem" "$stem"
    done
  fi
  printf '</Relationships>'
} > "$STAGE/customUI/_rels/customUI.xml.rels"

# package relationship: drop any existing extensibility rel, add exactly one
unzip -p "$BUILD" _rels/.rels > "$STAGE/_rels/.rels"
perl -pi -e 's{<Relationship[^>]*relationships/ui/extensibility[^>]*/>}{}g' "$STAGE/_rels/.rels"
perl -pi -e "s{</Relationships>}{<Relationship Id=\"rIdOBTRibbon\" Type=\"$REL06\" Target=\"customUI/customUI.xml\"/></Relationships>}" \
  "$STAGE/_rels/.rels"

# content types: a png Default must exist or the images resolve to nothing.
# The brackets in the part name are glob metacharacters to both unzip and zip,
# so every reference to it is either escaped or passed with -nw. Without that
# they match nothing and fail silently.
unzip -p "$BUILD" '\[Content_Types\].xml' > "$STAGE/Content_Types.xml"
[ -s "$STAGE/Content_Types.xml" ] || { echo "ERROR: could not read [Content_Types].xml from $INTO" >&2; exit 1; }
if ! grep -qi 'Extension="png"' "$STAGE/Content_Types.xml"; then
  perl -pi -e 's{<Types([^>]*)>}{<Types$1><Default Extension="png" ContentType="image/png"/>}' "$STAGE/Content_Types.xml"
  echo "note: added a png Default to [Content_Types].xml (the base workbook had none)"
fi
mv "$STAGE/Content_Types.xml" "$STAGE/[Content_Types].xml"

# swap the parts, leaving every other part of the base workbook byte-identical
zip -q -d "$BUILD" "customUI/*" "_rels/.rels" 2>/dev/null || true
zip -q -d -nw "$BUILD" "[Content_Types].xml" 2>/dev/null || true
# -D: no directory entries, matching how Excel writes the package
( cd "$STAGE" && zip -q -X -D -r "$BUILD" "_rels/.rels" "customUI" )
( cd "$STAGE" && zip -q -X -D -nw "$BUILD" "[Content_Types].xml" )

mv -f "$BUILD" "$ABS_OUT"

NIMG=$(find "$STAGE/customUI/images" -name '*.png' | wc -l | tr -d ' ')
echo "packed $DIR -> $OUT"
echo "  base    $INTO"
echo "  ribbon  customUI/customUI.xml (office/2006/01/customui)"
echo "  images  $NIMG png"
echo "Verify with: scripts/devtools/ribbon-doctor.sh $OUT"

#!/usr/bin/env bash
# ribbon-envelope.sh <workbook> [<workbook> ...]
#
# Moves a workbook's custom ribbon from the Office 2010 envelope
# (customUI/customUI14.xml + the office/2007 extensibility relationship,
# 2009/07 namespace) into the Office 2007 envelope
# (customUI/customUI.xml + the office/2006 relationship, 2006/01 namespace).
#
# Why: WPS Office rewrites a workbook's custom UI into the 2007 envelope but
# copies the XML body verbatim, leaving a 2009/07 namespace inside a 2006-typed
# part. Excel rejects that pairing and drops the ribbon without a word. Shipping
# in the 2007 envelope makes that rewrite a no-op. No OutbreakTools ribbon uses
# a 2010-only feature (backstage, contextMenus, tabSet), so nothing is lost.
#
# The control tree, callbacks and image relationships are carried across
# untouched; only the part names, the relationship type and the namespace move.
set -euo pipefail

NS14="http://schemas.microsoft.com/office/2009/07/customui"
NS12="http://schemas.microsoft.com/office/2006/01/customui"
REL07="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
REL06="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"

[ $# -gt 0 ] || { echo "usage: ribbon-envelope.sh <workbook> [...]" >&2; exit 2; }

for F in "$@"; do
  # capture first: `grep -q` closes the pipe early, which pipefail would
  # report as a failed unzip and turn into a false "skip"
  LIST=$(unzip -Z1 "$F" 2>/dev/null || true)
  if ! printf '%s\n' "$LIST" | grep -qx 'customUI/customUI14.xml'; then
    printf 'skip  %s (no customUI14 part)\n' "$F"; continue
  fi

  # resolve before entering the staging dir, so relative arguments still work
  ABS="$(cd "$(dirname "$F")" && pwd)/$(basename "$F")"

  STAGE=$(mktemp -d)
  trap 'rm -rf "$STAGE"' EXIT
  ( cd "$STAGE" && unzip -q "$ABS" \
      "customUI/customUI14.xml" "customUI/_rels/customUI14.xml.rels" "_rels/.rels" 2>/dev/null ) || true
  [ -f "$STAGE/customUI/customUI14.xml" ] || { echo "error: could not extract from $F" >&2; exit 1; }

  # 1. namespace, inside the part body
  perl -pi -e "s{\Q$NS14\E}{$NS12}g" "$STAGE/customUI/customUI14.xml"

  # 2. part names
  ENTRIES=( "_rels/.rels" "customUI/customUI.xml" )
  mv "$STAGE/customUI/customUI14.xml" "$STAGE/customUI/customUI.xml"
  if [ -f "$STAGE/customUI/_rels/customUI14.xml.rels" ]; then
    mv "$STAGE/customUI/_rels/customUI14.xml.rels" "$STAGE/customUI/_rels/customUI.xml.rels"
    ENTRIES+=( "customUI/_rels/customUI.xml.rels" )
  fi

  # 3. relationship type and target
  perl -pi -e "s{\Q$REL07\E}{$REL06}g; s{Target=\"customUI/customUI14\.xml\"}{Target=\"customUI/customUI.xml\"}g" \
    "$STAGE/_rels/.rels"

  # 4. swap the entries in the package, leaving every other part byte-identical
  zip -q -d "$F" "customUI/customUI14.xml" "customUI/_rels/customUI14.xml.rels" "_rels/.rels" 2>/dev/null || true
  ( cd "$STAGE" && zip -q -X "$ABS" "${ENTRIES[@]}" )

  rm -rf "$STAGE"; trap - EXIT
  printf 'moved %s -> 2007 envelope\n' "$F"
done

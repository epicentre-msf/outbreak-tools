#!/usr/bin/env bash
# ribbon-doctor.sh <workbook.xlsb|xlsm>
# Checks the package plumbing that makes a custom ribbon tab appear.
set -u
F="${1:?usage: ribbon-doctor.sh <workbook>}"
ok(){ printf '  ok    %s\n' "$1"; }
bad(){ printf '  FAIL  %s\n' "$1"; }
inf(){ printf '  --    %s\n' "$1"; }

echo "== $F"

unzip -t "$F" >/dev/null 2>&1 && ok "zip integrity" || bad "zip is damaged"

LIST=$(unzip -Z1 "$F" 2>/dev/null)

# 1. the extensibility relationship
RELS=$(unzip -p "$F" _rels/.rels 2>/dev/null)
[ -n "$RELS" ] || bad "_rels/.rels missing entirely"
EXT=$(printf '%s' "$RELS" | tr '<' '\n' | grep 'relationships/ui/extensibility')
if [ -z "$EXT" ]; then
  bad "NO ui/extensibility relationship in _rels/.rels  <-- ribbon can never load"
  printf '%s' "$LIST" | grep -qi customUI && \
    inf "yet customUI parts ARE in the zip -> the workbook was re-packaged/repaired"
else
  TARGET=$(printf '%s' "$EXT" | grep -oE 'Target="[^"]+"' | sed 's/Target="//;s/"//')
  RTYPE=$(printf '%s' "$EXT" | grep -oE 'office/20[0-9]{2}/relationships')
  ok "extensibility rel -> $TARGET  (type $RTYPE)"

  # 2. target exists, case-sensitively
  if printf '%s' "$LIST" | grep -qx "$TARGET"; then ok "target part present"
  else bad "target part '$TARGET' NOT in the zip (check case)"; exit 1; fi

  # 3. namespace <-> part name <-> rel type
  UI=$(unzip -p "$F" "$TARGET")
  NS=$(printf '%s' "$UI" | grep -oE 'schemas.microsoft.com/office/[0-9/]+/customui')
  inf "namespace $NS"
  case "$NS/$RTYPE" in
    */2009/07/customui/office/2007/relationships) ok "namespace matches rel type" ;;
    */2006/01/customui/office/2006/relationships) ok "namespace matches rel type" ;;
    *) bad "namespace and rel type do not pair -> Excel silently ignores the part" ;;
  esac
  case "$TARGET" in
    *customUI14.xml) [ "${NS##*office/}" = "2009/07/customui" ] || bad "customUI14.xml must use the 2009/07 namespace" ;;
    *customUI.xml)   [ "${NS##*office/}" = "2006/01/customui" ] || bad "2009/07 namespace in customUI.xml -> ignored; rename to customUI14.xml" ;;
  esac

  # 4. every custom image= id resolves
  DIR=$(dirname "$TARGET")
  IRELS=$(unzip -p "$F" "$DIR/_rels/$(basename "$TARGET").rels" 2>/dev/null)
  for id in $(printf '%s' "$UI" | grep -oE 'image="[^"]+"' | sed 's/image="//;s/"//' | sort -u); do
    line=$(printf '%s' "$IRELS" | tr '<' '\n' | grep "Id=\"$id\"")
    if [ -z "$line" ]; then bad "image=\"$id\" has NO relationship  <-- kills the whole ribbon"; continue; fi
    t=$(printf '%s' "$line" | grep -oE 'Target="[^"]+"' | sed 's/Target="//;s/"//')
    printf '%s' "$LIST" | grep -qx "$DIR/$t" \
      && ok "image $id -> $t" || bad "image \"$id\" points at $t which is not in the zip"
  done
fi

# 5. content types
CT=$(unzip -p "$F" '[Content_Types].xml' 2>/dev/null || unzip -p "$F" '\[Content_Types\].xml' 2>/dev/null)
printf '%s' "$CT" | grep -qi 'Extension="png"' && ok 'png Default in [Content_Types]' || bad 'no png Default -> images unresolvable'
printf '%s' "$CT" | grep -qi 'Extension="xml"' && ok 'xml Default in [Content_Types]' || bad 'no xml Default'

# 6. is the VBA still there (callbacks + a hint the file was format-converted)
printf '%s' "$LIST" | grep -qx 'xl/vbaProject.bin' \
  && ok "vbaProject.bin present" \
  || bad "no vbaProject.bin -> saved through a macro-free format; customUI usually dies with it"

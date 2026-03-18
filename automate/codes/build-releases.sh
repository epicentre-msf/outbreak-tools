#!/usr/bin/env bash
# Generates site/releases.qmd by scanning releases/old/
set -eu
cd "$(git rev-parse --show-toplevel)"

BASE="https://github.com/epicentre-msf/outbreak-tools/raw/dev/releases"
OUT="site/releases.qmd"

cat > "$OUT" <<'HEADER'
---
title: "Releases"
toc: false
---

Each archive contains a **designer**, a **setup**, and a **ribbon template** (`.xlsb`).

## Latest

| Branch | Download |
|--------|----------|
HEADER

# Latest main
if [ -f releases/latest/OBT-main-latest.zip ]; then
  echo "| Main (stable) | [OBT-main-latest.zip](${BASE}/latest/OBT-main-latest.zip){target=\"_blank\"} |" >> "$OUT"
fi

# Previous releases - Main
MAIN_ZIPS=$(ls releases/old/OBT-main-*.zip 2>/dev/null | xargs -n1 basename | sort -r || true)
if [ -n "$MAIN_ZIPS" ]; then
  cat >> "$OUT" <<'EOF'

## Previous releases

### Main

| Date | Download |
|------|----------|
EOF
  for z in $MAIN_ZIPS; do
    date=$(echo "$z" | grep -oE '[0-9]{4}-[0-9]{2}-[0-9]{2}')
    echo "| ${date} | [${z}](${BASE}/old/${z}){target=\"_blank\"} |" >> "$OUT"
  done
fi

# Previous releases - Dev
DEV_ZIPS=$(ls releases/old/OBT-dev-*.zip 2>/dev/null | xargs -n1 basename | sort -r || true)
if [ -n "$DEV_ZIPS" ]; then
  cat >> "$OUT" <<'EOF'

### Dev

| Date | Download |
|------|----------|
EOF
  for z in $DEV_ZIPS; do
    date=$(echo "$z" | grep -oE '[0-9]{4}-[0-9]{2}-[0-9]{2}')
    echo "| ${date} | [${z}](${BASE}/old/${z}){target=\"_blank\"} |" >> "$OUT"
  done
fi

echo "==> Generated $OUT with $(echo $MAIN_ZIPS $DEV_ZIPS | wc -w | tr -d ' ') releases"

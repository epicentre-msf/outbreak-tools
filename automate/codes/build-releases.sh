#!/usr/bin/env bash
#
# build-releases.sh — Generate site/releases.qmd from the GitHub Releases API.
# Replaces the old releases/old/ filesystem scan. Renders, per release, its changelog
# notes (the release body) and download links.
#
# Offline rendering test: RELEASES_JSON_FILE=fixture.json bash build-releases.sh
# (fixture = the JSON array returned by `gh api repos/<repo>/releases`).
#
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

REPO="${OBT_REPO:-epicentre-msf/outbreak-tools}"
OUT="${OUT:-site/releases.qmd}"

command -v jq >/dev/null 2>&1 || { echo "ERROR: jq not found." >&2; exit 1; }

if [ -n "${RELEASES_JSON_FILE:-}" ]; then
  rel="$(cat "$RELEASES_JSON_FILE")"
else
  command -v gh >/dev/null 2>&1 || { echo "ERROR: gh not found." >&2; exit 1; }
  rel="$(gh api "repos/$REPO/releases" --paginate)"
fi

# Drop the infra asset-store release everywhere; sort newest-first.
rel="$(printf '%s' "$rel" \
  | jq '[ .[] | select(.tag_name != "working-binaries") ] | sort_by(.published_at) | reverse')"

emit() { printf '%s\n' "$1" >> "$OUT"; }

# --- Header + Latest (static, version-independent download URLs) ---
cat > "$OUT" <<HEADER
---
title: "Releases"
toc: false
---

Each archive contains a **designer**, a **setup**, and a **ribbon template** (\`.xlsb\`).

## Latest

| Stream | Download |
|--------|----------|
| Main (stable) | [OBT-main-latest.zip](https://github.com/$REPO/releases/latest/download/OBT-main-latest.zip){target="_blank"} |
| Dev (pre-release) | [OBT-dev-latest.zip](https://github.com/$REPO/releases/download/dev-latest/OBT-dev-latest.zip){target="_blank"} |
HEADER

# --- Versioned releases (tag starts with 'v'): table + per-release notes ---
versioned="$(printf '%s' "$rel" | jq -c '[ .[] | select(.tag_name | test("^v")) ]')"
vcount="$(printf '%s' "$versioned" | jq 'length')"

if [ "$vcount" -gt 0 ]; then
  emit ""
  emit "## All releases"
  emit ""
  emit "| Version | Date | Stream | Download |"
  emit "|---------|------|--------|----------|"
  printf '%s' "$versioned" | jq -r '
    .[]
    | ( [ .assets[] | select(.name | test("^OBT-.*\\.zip$")) ] | .[0].browser_download_url // "" ) as $u
    | "| \(.tag_name) | \(.published_at[0:10]) | \(if .prerelease then "dev (pre)" else "main" end) | "
      + (if $u == "" then "—" else "[download](\($u)){target=\"_blank\"}" end) + " |"
  ' | awk '!seen[$0]++' >> "$OUT"

  emit ""
  emit "## Release notes"
  i=0
  while [ "$i" -lt "$vcount" ]; do
    name="$(printf '%s' "$versioned" | jq -r ".[$i].name // .[$i].tag_name")"
    date="$(printf '%s' "$versioned" | jq -r ".[$i].published_at[0:10]")"
    pre="$(printf '%s' "$versioned" | jq -r ".[$i].prerelease")"
    body="$(printf '%s' "$versioned" | jq -r ".[$i].body // \"\" | gsub(\"\r\";\"\")")"
    [ "$pre" = "true" ] && label="pre-release" || label="stable"
    emit ""
    emit "### $name ($label) — $date"
    emit ""
    if [ -n "$body" ]; then emit "$body"; else emit "_No notes._"; fi
    i=$((i + 1))
  done
fi

# --- Legacy archive (assets of the legacy-archive release) ---
legacy="$(printf '%s' "$rel" | jq -c '[ .[] | select(.tag_name == "legacy-archive") ] | .[0] // empty')"
if [ -n "$legacy" ] && [ "$(printf '%s' "$legacy" | jq '.assets | length')" -gt 0 ]; then
  emit ""
  emit "## Legacy archive"
  emit ""
  emit "Releases produced before the GitHub-Releases workflow."
  emit ""
  emit "| Date | Stream | Download |"
  emit "|------|--------|----------|"
  printf '%s' "$legacy" | jq -r '
    .assets[]
    | ( .name | capture("(?<date>[0-9]{4}-[0-9]{2}-[0-9]{2})").date // "" ) as $d
    | ( if (.name | test("-dev-")) then "dev" else "main" end ) as $s
    | "| \($d) | \($s) | [\(.name)](\(.browser_download_url)){target=\"_blank\"} |"
  ' | sort -r >> "$OUT"
fi

echo "==> Generated $OUT (versioned: $vcount; legacy: $([ -n "$legacy" ] && printf '%s' "$legacy" | jq '.assets|length' || echo 0))"

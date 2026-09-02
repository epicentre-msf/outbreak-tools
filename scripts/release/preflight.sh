#!/usr/bin/env bash
#
# preflight.sh [--check]
#
# Everything that must hold before working binaries leave this machine.
# push-assets.sh runs it first and pushes nothing if it fails.
#
#   1. envelope    no workbook carries a customUI14 part (conventions.md §11)
#   2. plumbing    ribbon-doctor.sh is clean on every workbook that has a ribbon
#   3. ribbons     the unpacked sources under ribbons/ match the workbooks
#   4. imports     hand-imported classes and modules match the repo
#   5. closure     the dev binaries carry every component their code names
#   6. trads       no translation the product asks for is missing from the
#                  designer binaries a build can read
#   7. rebuild     the ribbon sources are re-read out of the workbooks
#
# Cheap checks run first so an obvious failure does not cost an R startup, and
# the rebuild in step 7 runs only once every check is green — a push that is
# going to abort must leave the tree exactly as it found it.
#
# --check stops after step 6 and writes nothing.
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

CHECK_ONLY=0
case "${1:-}" in
  --check)   CHECK_ONLY=1 ;;
  -h|--help) sed -n '2,18p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
  "")        ;;
  *)         echo "unknown option: $1" >&2; exit 2 ;;
esac

# Colour like mock-import-drift.sh: on for a terminal, off for a pipe, a file
# or NO_COLOR, so redirected output stays greppable. The R scans below make
# the same choice on their own.
if [ -t 1 ] && [ -z "${NO_COLOR:-}" ]; then
  C_FAIL=$'\033[1;31m'; C_OK=$'\033[32m'; C_OFF=$'\033[0m'
else
  C_FAIL=""; C_OK=""; C_OFF=""
fi

DOCTOR="scripts/devtools/ribbon-doctor.sh"
DRIFT="scripts/devtools/mock-import-drift.sh"
SYNC="scripts/release/sync-ribbons.sh"
CLOSURE="scripts/devtools/workbook-closure-scan.R"
TRADBIN="scripts/devtools/translation-coverage-binaries.R"
for s in "$DOCTOR" "$DRIFT" "$SYNC" "$CLOSURE" "$TRADBIN"; do
  [ -f "$s" ] || { echo "ERROR: $s is missing." >&2; exit 1; }
done
command -v unzip >/dev/null 2>&1 || { echo "ERROR: unzip not found." >&2; exit 1; }
command -v Rscript >/dev/null 2>&1 || { echo "ERROR: Rscript not found." >&2; exit 1; }

# The workbooks the bundle ships and conventions.md §11 governs.
# src/tests/.input is bundled too and deliberately left out: those are frozen
# fixtures, and an old setup is *supposed* to still carry an old ribbon.
# ~$... is Excel's lock file for a workbook that is open right now. It is not a
# workbook, it is bundled all the same, and it must not be mistaken for one here.
WORKBOOKS=()
while IFS= read -r f; do [ -n "$f" ] && WORKBOOKS+=("$f"); done < <(
  ls .mock/*.xlsb src/bin/*/*.xlsb ribbons/_ribbontemplate_*.xlsb 2>/dev/null     | grep -v '/~\$' | sort -u
)
[ ${#WORKBOOKS[@]} -gt 0 ] || { echo "ERROR: no workbooks found (pull-assets.sh first?)" >&2; exit 1; }

TMPD="$(mktemp -d)"; trap 'rm -rf "$TMPD"' EXIT
FAILED=0

# --- 1. envelope -------------------------------------------------------------
# customUI14 is the 2010 envelope. WPS Office rewrites it into a 2006-typed part
# that keeps the 2009/07 namespace, Excel refuses that pairing, and the tab
# disappears with no error and no repair prompt. See conventions.md §11.
echo "==> customUI14 parts"
for wb in "${WORKBOOKS[@]}"; do
  hits=$(unzip -Z1 "$wb" 2>/dev/null | grep -i 'customUI14' || true)
  if [ -n "$hits" ]; then
    echo "  ${C_FAIL}FAIL${C_OFF}  $wb carries the 2010 envelope:" >&2
    printf '%s\n' "$hits" | sed 's/^/          /' >&2
    echo "        convert it: scripts/devtools/ribbon-envelope.sh $wb" >&2
    FAILED=1
  fi
done
[ $FAILED -eq 0 ] && echo "  ${C_OK}ok${C_OFF}    every ribbon is in the Office 2007 envelope"

# --- 2. ribbon plumbing ------------------------------------------------------
# ribbon-doctor.sh reports by printing; it does not signal with its exit status,
# so the FAIL lines are what is read here.
echo "==> ribbon plumbing (ribbon-doctor)"
for wb in "${WORKBOOKS[@]}"; do
  if ! unzip -Z1 "$wb" 2>/dev/null | grep -q '^customUI/'; then
    echo "  --    $wb carries no ribbon, skipped"
    continue
  fi
  bash "$DOCTOR" "$wb" >"$TMPD/doctor.out" 2>&1 || true
  if grep -q '^  FAIL' "$TMPD/doctor.out"; then
    echo "  ${C_FAIL}FAIL${C_OFF}  $wb" >&2
    sed 's/^/        /' "$TMPD/doctor.out" >&2
    FAILED=1
  else
    echo "  ${C_OK}ok${C_OFF}    $wb"
  fi
done

# --- 3. ribbon sources -------------------------------------------------------
echo
bash "$SYNC" --check || FAILED=1

# --- 4. hand-imported VBA ----------------------------------------------------
# Classes and modules pasted into the mock workbooks by hand travel inside the
# .xlsb while the .cls and .bas they came from keep moving in git.
# The drift scan picks its own colours the same way this script does; forcing
# them off here hid the red rows a reader is meant to catch at a glance.
echo
echo "==> hand-imported components (mock-import-drift)"
bash "$DRIFT" || FAILED=1

# --- 5. component closure ----------------------------------------------------
# The dev binaries compile from whatever was pasted into them, and a component
# left off the paste is a project-wide compile failure in the field. The scan
# reads the pasted code itself out of vbaProject.bin, in plain R on either OS.
# The repo .Rprofile loads packages these scans never ask for, and R dying at
# startup over them reads exactly like a failed check, so it is held back.
echo
echo "==> component closure (workbook-closure-scan)"
R_PROFILE_USER=/dev/null Rscript "$CLOSURE" || FAILED=1

# --- 6. designer translations ------------------------------------------------
# A generated linelist reads its translation tables from the designer BINARY,
# not from the trads master, so a tag added to the master and never pasted into
# the designer still ships a screen reading the tag itself. Gate is missing
# tags only: dead rows are upkeep, not a broken screen.
echo
echo "==> designer translations (translation-coverage-binaries)"
R_PROFILE_USER=/dev/null Rscript "$TRADBIN" || FAILED=1

if [ $FAILED -ne 0 ]; then
  echo >&2
  echo "${C_FAIL}ERROR: preflight failed — nothing was pushed and nothing was rewritten.${C_OFF}" >&2
  exit 1
fi

[ $CHECK_ONLY -eq 1 ] && { echo; echo "${C_OK}Preflight is clean.${C_OFF}"; exit 0; }

# --- 7. rebuild --------------------------------------------------------------
echo
bash "$SYNC"
echo
echo "${C_OK}Preflight is clean.${C_OFF}"

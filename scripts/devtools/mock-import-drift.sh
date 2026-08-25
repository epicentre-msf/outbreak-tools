#!/usr/bin/env bash
#
# mock-import-drift.sh -- does a hand-imported class or module inside a mock
# workbook still match the source file it was pasted from?
#
#   scripts/devtools/mock-import-drift.sh [target ...] [options]
#
#   targets:  setup | designer | unit_test | mastersetup | all   (default: all)
#
#   options:
#     --fast             skip the code comparison; ask git history instead
#     --since <ISO8601>  --fast only: compare against this instant, not the push
#     --repo  <owner/nm> --fast only: the GitHub repo holding the asset store
#     --list             just list what is inside each workbook, check nothing
#     --no-color         plain output (also honours NO_COLOR and a non-tty)
#     -h, --help         this text
#
# Every component is listed one per line, alphabetically. The ones that differ
# from the repo are in red; a component with nothing to compare against is
# dimmed.
#
# WHY THIS EXISTS
# -----------------------------------------------------------------------------
# Some classes and modules are not imported by any build step. They are pasted
# into the mock workbooks by hand and then travel inside the .xlsb, which is
# pushed off-git by scripts/release/push-assets.sh. Meanwhile the .cls and .bas
# they were pasted from keep moving in git, because a bug gets fixed in one of
# them like any other file. Nothing connects the two, so a binary can quietly
# carry a stale copy of a class that was fixed weeks ago.
#
# THE TWO MODES
# -----------------------------------------------------------------------------
# DEFAULT reads the VBA source back out of the workbook and compares it with the
# file in the repo, through scripts/devtools/vba-inspect.R. It answers "this
# copy differs", which needs no reference date and cannot be fooled: an edit
# made in the VBE and never exported shows up as clearly as a commit that was
# never imported.
#
# --fast skips all of that and asks git whether the source file has moved since
# the working binaries were last pushed. It is the older, cheaper check and it
# is wrong in both directions -- a comment-only commit reads as drift, and
# anything that fell out of sync BEFORE the last push is invisible to it,
# because `git log --since` only ever looks forward from that instant.
#
# THE FOUR ANSWERS
# -----------------------------------------------------------------------------
#   same       the code matches, character for character
#   cosmetic   it matches once the VBE's own rewrites are undone -- identifier
#              capitalisation and whitespace between tokens. Nothing to do.
#              vba-inspect.R explains which rewrites are covered.
#   differs    a real difference. Re-import it.
#   (dimmed)   nothing to compare against: no file in the repo, or the source
#              could not be read back out of the workbook.
#
# THE WORKBOOK IS NEVER OPENED IN PLACE. Every target is copied into a fresh
# mktemp -d and unzipped there, and the temp tree is removed on exit. The
# originals are the off-git source of truth and one stray write loses work that
# is in no commit. A workbook open in Excel is safe to run this against.
# =============================================================================
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

INSPECT="scripts/devtools/vba-inspect.R"
REPO="${OBT_REPO:-epicentre-msf/outbreak-tools}"
TAG="working-binaries"
ASSET="working-binaries.tar.gz"
# User forms are hand-maintained too and their .frm/.frx sit outside src/, under
# .mock/forms/<product>/. Imports.frm and the designer's F_* forms are found
# there and nowhere else, so this root is not optional.
SEARCH_ROOTS=(src/classes src/modules src/tests .mock/forms)

SINCE=""
SINCE_LABEL="working binaries last pushed"
LIST_ONLY=0
FAST=0
USE_COLOR="auto"
TARGETS=()

workbook_for() {
  case "$1" in
    setup)       echo ".mock/setup_mock.xlsb" ;;
    designer)    echo ".mock/designer_mock.xlsb" ;;
    unit_test)   echo ".mock/unit_test_mock.xlsb" ;;
    mastersetup) echo ".mock/mastersetup_mock.xlsb" ;;
    *)           echo "" ;;
  esac
}

# The header doubles as the help text. It is printed from line 2 down to the
# WHY THIS EXISTS marker, so adding notes below that never lengthens --help.
usage() {
  awk 'NR>1 && /^# WHY THIS EXISTS/{exit} NR>1 && /^#/{sub(/^# ?/,""); print}' "$0"
  exit "${1:-0}"
}

while [ $# -gt 0 ]; do
  case "$1" in
    --fast)  FAST=1; shift ;;
    --since) SINCE="${2:?--since needs an ISO8601 instant}"
             SINCE_LABEL="comparing against (--since)"; shift 2 ;;
    --repo)  REPO="${2:?--repo needs owner/name}"; shift 2 ;;
    --list)  LIST_ONLY=1; shift ;;
    --no-color) USE_COLOR="never"; shift ;;
    -h|--help) usage 0 ;;
    -*) echo "ERROR: unknown option $1" >&2; usage 1 ;;
    all) TARGETS+=(setup designer unit_test mastersetup); shift ;;
    *)
      if [ -z "$(workbook_for "$1")" ]; then
        echo "ERROR: unknown target '$1' (setup|designer|unit_test|mastersetup|all)" >&2
        exit 1
      fi
      TARGETS+=("$1"); shift ;;
  esac
done
[ ${#TARGETS[@]} -gt 0 ] || TARGETS=(setup designer unit_test mastersetup)

command -v Rscript >/dev/null 2>&1 || {
  echo "ERROR: Rscript not found; this tool reads the workbooks with R." >&2; exit 1; }
[ -f "$INSPECT" ] || { echo "ERROR: $INSPECT is missing." >&2; exit 1; }

if [ -n "$SINCE" ] && [ "$FAST" -eq 0 ]; then
  echo "NOTE: --since only applies to --fast; the default compares code, not dates." >&2
fi

# --- the instant to compare against, for --fast only -------------------------
# The asset store is written with `gh release upload --clobber`, which replaces
# the asset, so the asset's updatedAt IS the moment push-assets.sh last ran.
if [ "$FAST" -eq 1 ] && [ "$LIST_ONLY" -eq 0 ] && [ -z "$SINCE" ]; then
  if ! command -v gh >/dev/null 2>&1; then
    echo "ERROR: gh CLI not found, so the last push time cannot be read." >&2
    echo "       Install it, or pass the instant yourself: --since 2026-08-15T11:17:55Z" >&2
    exit 1
  fi
  SINCE="$(gh release view "$TAG" -R "$REPO" --json assets \
             --jq ".assets[] | select(.name==\"$ASSET\") | .updatedAt" 2>/dev/null || true)"
  if [ -z "$SINCE" ]; then
    echo "ERROR: could not read the '$ASSET' timestamp from release '$TAG' in $REPO." >&2
    echo "       Check 'gh auth status', or pass --since <ISO8601>." >&2
    exit 1
  fi
fi

# --- read one workbook -------------------------------------------------------
# Emits the TSV rows of vba-inspect.R: kind, name, path, status, detail.
# The copy and the unzip happen in a temp dir that is removed before returning,
# so the workbook itself is only ever read.
analyze() {
  local wb="$1" mode="$2" tmpd rc=0
  tmpd="$(mktemp -d)"

  cp "$wb" "$tmpd/book.xlsb"
  if ! unzip -o -q "$tmpd/book.xlsb" -d "$tmpd/unz" 2>/dev/null; then
    echo "ERROR: $wb is not readable as a zip package." >&2
    rm -rf "$tmpd"; return 1
  fi
  if [ ! -f "$tmpd/unz/xl/vbaProject.bin" ]; then
    echo "ERROR: $wb holds no xl/vbaProject.bin (no VBA project)." >&2
    rm -rf "$tmpd"; return 1
  fi

  Rscript "$INSPECT" "$tmpd/unz/xl/vbaProject.bin" "$mode" "${SEARCH_ROOTS[@]}" || rc=$?

  rm -rf "$tmpd"
  return $rc
}

# GNU stat is tried first and its success is tested on the value, not on the
# pipeline: GNU stat reads -f as 'filesystem status', so probing the BSD form
# first succeeds everywhere and prints a filesystem dump instead of a date.
file_time() {
  local t
  t=$(stat -c '%y' "$1" 2>/dev/null) && { printf '%s' "${t:0:16}"; return 0; }
  stat -f '%Sm' -t '%Y-%m-%d %H:%M' "$1" 2>/dev/null
}

# --- report ------------------------------------------------------------------
# Colour is on for a terminal and off for a pipe, a file or NO_COLOR, so the
# output stays greppable when it is redirected.
if [ -t 1 ] && [ -z "${NO_COLOR:-}" ] && [ "$USE_COLOR" != "never" ]; then
  C_DRIFT=$'\033[1;31m'
  C_DIM=$'\033[2m'
  C_HEAD=$'\033[1m'
  C_OFF=$'\033[0m'
else
  C_DRIFT=""; C_DIM=""; C_HEAD=""; C_OFF=""
fi

drift_found=0
missing_wb=0
TAB="$(printf '\t')"

if [ "$LIST_ONLY" -eq 0 ]; then
  if [ "$FAST" -eq 1 ]; then
    printf '%s: %s\n' "$SINCE_LABEL" "$SINCE"
    [ "$SINCE_LABEL" = "working binaries last pushed" ] && printf 'repo: %s\n' "$REPO"
    printf '%sfast mode: comparing dates, not code%s\n' "$C_DIM" "$C_OFF"
  else
    printf 'comparing the code inside each workbook with the repo\n'
  fi
  echo
fi

for t in "${TARGETS[@]}"; do
  wb="$(workbook_for "$t")"

  if [ ! -f "$wb" ]; then
    printf '%b== %s%b  %s(%s)%s\n' "$C_HEAD" "$t" "$C_OFF" "$C_DIM" "$wb" "$C_OFF"
    printf '   workbook not found, skipped\n\n'
    missing_wb=1
    continue
  fi

  printf '%b== %s%b  %s(%s, saved %s)%s\n' \
         "$C_HEAD" "$t" "$C_OFF" "$C_DIM" "$wb" "$(file_time "$wb")" "$C_OFF"

  if [ "$LIST_ONLY" -eq 1 ] || [ "$FAST" -eq 1 ]; then mode="list"; else mode="content"; fi
  parsed="$(analyze "$wb" "$mode")" || { echo; continue; }

  rows=()
  n_drift=0; n_clean=0; n_cos=0; n_orphan=0; n_sheet=0

  while IFS="$TAB" read -r kind name path status detail; do
    [ -n "$name" ] || continue

    # A Document is the workbook or one of its sheets. Its code lives in the
    # .xlsb and there is no file anywhere to compare it against.
    if [ "$kind" = "Document" ]; then
      n_sheet=$((n_sheet + 1))
      [ "$LIST_ONLY" -eq 1 ] && rows+=("$(printf '3\t%s\tsheet module\t' "$name")")
      continue
    fi

    if [ "$LIST_ONLY" -eq 1 ]; then
      rows+=("$(printf '0\t%s\t\t' "$name")")
      continue
    fi

    if [ "$FAST" -eq 0 ]; then
      case "$status" in
        differs)
          rows+=("$(printf '2\t%s\t%s\t%s' "$name" "$path" "$detail")")
          n_drift=$((n_drift + 1)) ;;
        same)
          rows+=("$(printf '0\t%s\t%s\t%s' "$name" "$path" "$detail")")
          n_clean=$((n_clean + 1)) ;;
        cosmetic)
          rows+=("$(printf '0\t%s\t%s\t%s' "$name" "$path" "$detail")")
          n_cos=$((n_cos + 1)) ;;
        *)
          # nosource, noextract and empty all mean the same thing for the
          # verdict: there is nothing here to act on.
          rows+=("$(printf '1\t%s\t%s\t%s' "$name" "$path" \
                    "$([ "$status" = "nosource" ] && echo "no source in repo" || echo "$detail")")")
          n_orphan=$((n_orphan + 1)) ;;
      esac
      continue
    fi

    # --fast: git history on the resolved source file(s)
    srcs=()
    while IFS= read -r line; do
      [ -n "$line" ] && srcs+=("$line")
    done < <(find "${SEARCH_ROOTS[@]}" -type f \
                  \( -name "$name.cls" -o -name "$name.bas" \
                     -o -name "$name.frm" -o -name "$name.frx" \) 2>/dev/null | sort)

    if [ ${#srcs[@]} -eq 0 ]; then
      rows+=("$(printf '1\t%s\t-\tno source in repo' "$name")")
      n_orphan=$((n_orphan + 1))
      continue
    fi

    primary=""
    for f in "${srcs[@]}"; do
      case "$f" in *.frx) ;; *) [ -n "$primary" ] || primary="$f" ;; esac
    done
    [ -n "$primary" ] || primary="${srcs[0]}"

    commits=0; dirty=""; extra=""
    for f in "${srcs[@]}"; do
      n="$(git log --since="$SINCE" --pretty=%h -- "$f" | wc -l | tr -d ' ')"
      commits=$((commits + n))
      [ -n "$(git status --porcelain -- "$f")" ] && dirty="yes"
      case "$f" in *.frx) [ "$n" -gt 0 ] && extra=" (layout too)" ;; esac
    done

    if [ -n "$dirty" ]; then
      note="uncommitted"
      [ "$commits" -gt 0 ] && note="uncommitted + $commits commit(s)"
      rows+=("$(printf '2\t%s\t%s\t%s%s' "$name" "$primary" "$note" "$extra")")
      n_drift=$((n_drift + 1))
    elif [ "$commits" -gt 0 ]; then
      rows+=("$(printf '2\t%s\t%s\t%s commit(s) since%s' "$name" "$primary" "$commits" "$extra")")
      n_drift=$((n_drift + 1))
    else
      rows+=("$(printf '0\t%s\t%s\t' "$name" "$primary")")
      n_clean=$((n_clean + 1))
    fi
  done <<< "$parsed"

  if [ ${#rows[@]} -gt 0 ]; then
    printf '%s\n' "${rows[@]}" | sort -t"$TAB" -k2,2f | \
    while IFS="$TAB" read -r rank name path note; do
      case "$rank" in
        2) printf '   %b%-26s%b %-46s %b%s%b\n' \
                  "$C_DRIFT" "$name" "$C_OFF" "$path" "$C_DRIFT" "$note" "$C_OFF" ;;
        1) printf '   %b%-26s%b %-46s %b%s%b\n' \
                  "$C_DIM" "$name" "$C_OFF" "$path" "$C_DIM" "$note" "$C_OFF" ;;
        3) printf '   %b%-26s %s%b\n' "$C_DIM" "$name" "$path" "$C_OFF" ;;
        *) if [ "$LIST_ONLY" -eq 1 ]; then printf '   %s\n' "$name"
           else printf '   %-26s %-46s %s%s%s\n' "$name" "$path" "$C_DIM" "$note" "$C_OFF"; fi ;;
      esac
    done
  fi

  [ "$LIST_ONLY" -eq 1 ] && { echo; continue; }

  [ "$n_drift" -gt 0 ] && drift_found=1

  printf '   %s--%s\n' "$C_DIM" "$C_OFF"
  if [ "$n_drift" -gt 0 ]; then
    printf '   %b%d to re-import%b, %d up to date' "$C_DRIFT" "$n_drift" "$C_OFF" "$n_clean"
  else
    printf '   0 to re-import, %d up to date' "$n_clean"
  fi
  [ "$n_cos" -gt 0 ] && printf ' (+%d matching after the VBE rewrites)' "$n_cos"
  [ "$n_orphan" -gt 0 ] && printf ', %d not comparable' "$n_orphan"
  [ "$n_sheet" -gt 0 ] && printf ', %d sheet module(s) skipped' "$n_sheet"
  printf '\n\n'
done

if [ "$LIST_ONLY" -eq 1 ]; then exit 0; fi

if [ "$drift_found" -eq 1 ]; then
  echo "Drift found. Re-import the components listed above in the VBE, then run"
  echo "scripts/release/push-assets.sh so the store matches the source again."
  exit 1
fi

[ "$missing_wb" -eq 1 ] && exit 1
echo "Every hand-imported component matches the repo."
exit 0

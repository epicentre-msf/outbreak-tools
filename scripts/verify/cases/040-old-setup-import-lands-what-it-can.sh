#!/usr/bin/env bash
# title: Importing an old setup reports its faults and still lands what it can
# asserts: a clean old setup converts whole; a faulty one reports the fault, keeps its other sheets and still builds
#
# WHAT THIS CASE IS ABOUT
# -----------------------------------------------------------------------------
# A user arrives with a setup they filled in months ago. Converting it means
# filling a fresh copy of the CURRENT empty setup from it, and an old file can
# carry anything: a sheet whose table moved, a translation column that cannot
# be read, a formula nobody fixed. The requirement is that such a setup does
# not take the whole import down with it -- what can land, lands, and what
# cannot is reported.
#
# WHAT IT ASSERTS, AND THE GAP IT DOES NOT
# -----------------------------------------------------------------------------
# This case pins TODAY'S behaviour so a regression shows, which is a narrower
# claim than the requirement. Measured 2026-11-25 over the five old setups:
#
#   ebola-2026-11-19       OK, 381 dictionary rows
#   cholera-nga-borno      OK, 143
#   measles-training-2024  OK, 106
#   generic                OK,  99
#   malaria-aweil          ERROR 91 on its TRANSLATIONS sheet, 63 rows landed
#
# For malaria-aweil the four sheets BEFORE translations all landed --
# dictionary 63, choices 401, exports, analysis -- and the import then raised.
# So data does survive, the fault is reported, and the target setup still
# generates a linelist.
#
# THE GAP, written down so nobody reads this pass as the requirement being met:
# the import RAISES on the first sheet that fails instead of collecting a fault
# per sheet and carrying on. Translations is last in the list the import asks
# for, which is the only reason the other four landed. A setup whose
# DICTIONARY failed would lose the four sheets after it, and nothing in this
# case or in the product would say which ones were skipped. Making the import
# collect per-sheet faults is the change the requirement actually asks for, and
# when it lands this case should grow an assertion that a fault in an EARLY
# sheet still leaves the later ones imported.

set -uo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/../lib/verify-lib.sh"

need_command Rscript

# THE CONVERSION DRIVER. Nothing in scripts/ drives
# HeadlessBuild.ImportSetupFromWorkbook: its only caller is
# TestHeadlessSetupFill, which carries its paths as compiled constants. The
# driver lives in the headless draft folder until it is promoted into
# scripts/headless/windows/, so both homes are tried.
CONVERT=""
for candidate in "$REPO/scripts/headless/windows/convert-setup.ps1" \
                 "$FIXTURES/convert-setup.ps1"
do
    if [ -f "$candidate" ]; then CONVERT="$candidate"; break; fi
done
need_file "$CONVERT" "the conversion driver"

need_file "$EMPTY_SETUP" "the current empty setup"

OLD_DIR="$FIXTURES/setups/old"
[ -d "$OLD_DIR" ] || skip_case "no old setups under $OLD_DIR"

win() { cygpath -w "$1" 2>/dev/null || printf '%s' "$1"; }

OUT="$VERIFY_WORK/040-import"
mkdir -p "$OUT"

# convert_one <source> -> prints "<exit> <rows> <outcome>"
convert_one() {
    local src="$1"
    local stem; stem="$(basename "$src")"; stem="${stem%.*}"
    local target="$OUT/$stem.${EMPTY_SETUP##*.}"
    local log="$OUT/convert-$stem.log"
    shift

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(win "$CONVERT")" \
        -Source "$(win "$src")" -Target "$(win "$target")" \
        -Template "$(win "$EMPTY_SETUP")" "$@" \
        > "$log" 2>&1
    local code=$?
    local rows outcome
    rows="$(sed -n 's/^rows=//p' "$log" | head -1)"
    outcome="$(grep -E '^(OK|ERROR )' "$log" | tail -1)"
    printf '%s|%s|%s|%s\n' "$code" "${rows:-?}" "${outcome:-none}" "$target"
}

# --- 1. A CLEAN OLD SETUP CONVERTS WHOLE ------------------------------------
CLEAN_SRC=""
for candidate in "$OLD_DIR/ebola-2026-11-19.xlsb" \
                 "$OLD_DIR/cholera-nga-borno.xlsb" \
                 "$OLD_DIR/measles-training-2024.xlsb" \
                 "$OLD_DIR/generic.xlsb"
do
    if [ -f "$candidate" ]; then CLEAN_SRC="$candidate"; break; fi
done

if [ -n "$CLEAN_SRC" ]; then
    step "converting $(basename "$CLEAN_SRC"), expected to convert whole"
    IFS='|' read -r code rows outcome target <<< "$(convert_one "$CLEAN_SRC")"
    say "exit $code, rows $rows, outcome: $outcome"
    assert_equal "$code" "0" "a clean old setup converts and the driver exits 0"
    assert_contains "$OUT/convert-$(basename "${CLEAN_SRC%.*}").log" "^OK" "the import answered OK"
    assert_at_least "$rows" 50 "the dictionary came across (>=50 rows)"
    assert_file "$target" "the converted setup was written"
else
    say "no clean old setup on this machine; the whole-conversion half is skipped"
fi

# --- 2. A FAULTY OLD SETUP REPORTS AND STILL LANDS WHAT IT CAN --------------
FAULTY_SRC="$OLD_DIR/malaria-aweil.xlsb"
if [ ! -f "$FAULTY_SRC" ]; then
    say "no faulty old setup on this machine; the resilience half is skipped"
    check_case "a clean old setup converts whole"
fi

step "converting $(basename "$FAULTY_SRC"), expected to report a fault"
IFS='|' read -r fcode frows foutcome ftarget <<< "$(convert_one "$FAULTY_SRC")"
FAULTY_LOG="$OUT/convert-malaria-aweil.log"
say "exit $fcode, rows $frows, outcome: $foutcome"

# THE FAULT IS REPORTED, in words and with a number. A silent partial import is
# the failure this guards: it would hand back a setup missing a sheet and say
# nothing.
assert_contains "$FAULTY_LOG" "ERROR" "the faulty setup's import reports an error rather than passing silently"

# THE DATA THAT COULD LAND, LANDED. 63 is what the dictionary of this setup
# holds; the assertion is that it is there at all, so the number is a floor
# rather than an equality.
assert_at_least "$frows" 1 "the dictionary landed despite the fault"
say "the import landed $frows dictionary row(s) before the fault"

# THE FILE EXISTS, so the user has something rather than nothing.
assert_file "$ftarget" "a setup was still written for the user"

# --- 3. AND THE PARTLY-CONVERTED SETUP IS STILL USABLE ----------------------
# The point of landing what it can: a linelist can be built from the result.
# This is the assertion that makes "it did not break the import" mean
# something to a user rather than only to a log reader.
if [ -f "$ftarget" ] && [ -f "$DESIGNER" ] && [ -f "$TEMPLATE" ]; then
    build_linelist "$OUT" "ll-partial" "$ftarget" --lllang=ENG-English
    say "build of the partly-converted setup exited $?"
    assert_file "$OUT/ll-partial.xlsm" "a linelist builds from the partly-converted setup"

    # Its own generation errors are the setup's, not the import's: this setup
    # has case_id_auto written as IF(ISBLANK(case_id, "", ... which closes
    # ISBLANK in the wrong place. Printed rather than asserted, because they
    # belong to the source file.
    if [ -f "$OUT/ll-partial-generation.txt" ]; then
        n=$(grep -c "\[Error\]" "$OUT/ll-partial-generation.txt" 2>/dev/null || echo 0)
        say "the generation logged $n error(s), which are this setup's own formula faults"
    fi
else
    say "no designer or template, so the usability half is skipped"
fi

say ""
say "GAP, by design of this case: the import raises on the FIRST failing sheet"
say "rather than collecting a fault per sheet. Translations is last in the list,"
say "which is the only reason the other four sheets landed here."

check_case "an old setup's import reports its faults and lands what it can"

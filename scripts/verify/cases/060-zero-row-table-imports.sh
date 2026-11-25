#!/usr/bin/env bash
# title: A source table with zero data rows imports without raising
# asserts: a setup whose Tab_Translations has no data body answers OK, lands its other sheets, and the real aweil setup converts
#
# WHAT THIS CASE IS ABOUT
# -----------------------------------------------------------------------------
# A table whose ROWS were deleted has no data body at all: Excel keeps the
# header and a phantom insert row, and DataBodyRange answers Nothing. That is
# the shape `setup_20251201_aweil 1.xlsb` ships in -- Tab_Translations at B4:E5
# with headers for English, Francais, Espanol and Portugese and no translation
# rows, because nobody ever translated that setup.
#
# Importing it raised `ERROR 91: Object variable or With block variable not
# set` and aborted the whole conversion, losing a setup a user would otherwise
# have got back.
#
# THE SHAPE CASE 050 DOES NOT COVER
# -----------------------------------------------------------------------------
# 050 empties a sheet with ClearContents, which blanks the cells and leaves the
# table its height, so DataBodyRange stays a real range and all five sheets
# answer OK. This case deletes the ROWS. The two look alike to a reader and
# behave differently, which is the whole reason both exist.
#
# THE ROOT CAUSE, found 2026-11-25
# -----------------------------------------------------------------------------
# `CustomTable.ImportAll` read `impTab.DataRange().Value` with no guard.
# `CustomTable.DataRange` leaves its range unset when the source ListObject's
# DataBodyRange is Nothing, so `.Value` on Nothing raised 91 and
# `RollbackImport` re-raised it out of the import.
#
# Only the Translations sheet reached that code: `SetupImport.Import` passes
# `keepSourceHeaders:=True` for Translations alone, and that flag is the only
# route into `ImportAll`. Every other sheet goes through `ImportObject`, which
# already guarded `sourceRng Is Nothing` and carries a comment recording an
# earlier fix for this very error on an empty table. The guard was never
# applied to the other branch.
#
# Proved by construction before the fix: deleting the Translations rows of
# `ebola-2026-11-19.xlsb`, which imports whole at 381 dictionary rows,
# reproduced ERROR 91 exactly.
#
# WHAT THIS CASE DOES NOT CLAIM
# -----------------------------------------------------------------------------
# That an empty translation table is a good thing to ship. The target comes out
# with no translations, and a setup with no translations is a setup a user still
# has to translate. The requirement is that the import survives it and hands
# back what it could.
#
# It also tests ONE table's emptiness at a time. Two empty tables in one setup
# is a shape nobody has asked about.

set -uo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/../lib/verify-lib.sh"

CASE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
FIXTURE_PS1="$CASE_DIR/060-zero-row-table.ps1"

CONVERT="$(first_existing \
    "$REPO/scripts/headless/windows/convert-setup.ps1" \
    "$FIXTURES/convert-setup.ps1" || true)"

need_file "$CONVERT" "the conversion driver"
need_file "$FIXTURE_PS1" "the fixture builder"
need_file "$EMPTY_SETUP" "the current empty setup"

# A CLEAN old setup, so the emptied table is the only thing wrong with it.
SOURCE="$(first_existing \
    "$FIXTURES/setups/old/ebola-2026-11-19.xlsb" \
    "$FIXTURES/setups/old/cholera-nga-borno.xlsb" \
    "$FIXTURES/setups/old/measles-training-2024.xlsb" || true)"
need_file "$SOURCE" "a clean old setup to empty a table of"

OUT="$VERIFY_WORK/060-zero-row"
mkdir -p "$OUT"

win() { cygpath -w "$1" 2>/dev/null || printf '%s' "$1"; }

convert() {
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(win "$CONVERT")" \
        -Source "$(win "$1")" -Target "$(win "$2")" -Template "$(win "$EMPTY_SETUP")" \
        > "$3" 2>&1
    return 0
}

outcome_of() { grep -E '^(OK|ERROR )' "$1" 2>/dev/null | tr -d '\r' | tail -1; }
rows_of()    { sed -n 's/^rows=//p' "$1" 2>/dev/null | tr -d '\r' | head -1; }

# =============================================================================
# THE BUILT FIXTURE -- a clean setup with its translation rows deleted
# =============================================================================
step "deleting the data rows of Tab_Translations in $(basename "$SOURCE")"

FIXTURE="$OUT/src-zero-row-trads.${SOURCE##*.}"
FLOG="$OUT/fixture.log"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(win "$FIXTURE_PS1")" \
    -Source "$(win "$SOURCE")" -Target "$(win "$FIXTURE")" \
    -Sheet "Translations" -Table "Tab_Translations" \
    > "$FLOG" 2>&1
FIXTURE_EXIT=$?
sed 's/^/      /' "$FLOG"

if [ "$FIXTURE_EXIT" -ne 0 ]; then
    skip_case "the fixture could not be made: $(grep -m1 '^ERROR' "$FLOG" | tr -d '\r')"
fi

assert_contains "$FLOG" "BODY_AFTER=Nothing" "the fixture's table has no data body"
say "the table held $(sed -n 's/^ROWS_BEFORE=//p' "$FLOG" | tr -d '\r' | head -1) row(s) before"

CLOG="$OUT/convert-fixture.log"
convert "$FIXTURE" "$OUT/conv-fixture.${EMPTY_SETUP##*.}" "$CLOG"

OUTCOME="$(outcome_of "$CLOG")"
ROWS="$(rows_of "$CLOG")"
say "outcome: ${OUTCOME:-nothing}   dictionary rows: ${ROWS:-?}"

# THE REQUIREMENT. An empty source table must not raise. Error 91 in
# particular is the regression this case exists for, so it is named.
assert_equal "$OUTCOME" "OK" "a table with no data body imports without raising"
assert_not_contains "$CLOG" "ERROR 91" "no 'Object variable or With block variable not set'"
assert_file "$OUT/conv-fixture.${EMPTY_SETUP##*.}" "a setup was written"

# AND THE OTHER SHEETS STILL LANDED. The fault used to abort the whole import,
# so the dictionary surviving is what says it no longer does.
assert_at_least "${ROWS:-0}" 100 "the dictionary landed despite the empty translation table"

# =============================================================================
# THE REAL SETUP that found this
# =============================================================================
# The fixture above reproduces the shape; this is the file that ships in it.
# Skipped rather than failed when it is not on the machine, like every other
# fixture in this suite.
AWEIL="$FIXTURES/setups/old/malaria-aweil.xlsb"
if [ -f "$AWEIL" ]; then
    say ""
    step "the setup that found this: $(basename "$AWEIL")"
    ALOG="$OUT/convert-aweil.log"
    convert "$AWEIL" "$OUT/conv-aweil.${EMPTY_SETUP##*.}" "$ALOG"

    A_OUTCOME="$(outcome_of "$ALOG")"
    A_ROWS="$(rows_of "$ALOG")"
    say "outcome: ${A_OUTCOME:-nothing}   dictionary rows: ${A_ROWS:-?}"

    assert_equal "$A_OUTCOME" "OK" "the aweil setup converts whole"
    assert_not_contains "$ALOG" "ERROR 91" "the aweil setup raises no 91"
    assert_at_least "${A_ROWS:-0}" 50 "the aweil dictionary landed its rows"
else
    say ""
    say "NOT COVERED: $(basename "$AWEIL") is not on this machine, so the real file was not tried"
fi

left="$(wait_for_no_excel 20)"
assert_equal "$left" "0" "no Excel was left running after the conversions"

check_case "a source table with no data body imports without raising"

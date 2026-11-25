#!/usr/bin/env bash
# title: An empty worksheet in a setup being imported does not break the import
# asserts: each imported sheet, emptied in turn, answers OK or a reported error, still writes a setup, and leaves the other sheets' data intact
#
# WHAT THIS CASE IS ABOUT
# -----------------------------------------------------------------------------
# A user converts a setup they filled months ago, and one of its worksheets
# holds nothing: they deleted their rows, or they never filled that sheet in.
# The import asks for five sheets by name -- Dictionary, Choices, Exports,
# Analysis, Translations -- and the requirement is that an empty one of them
# does not take the operation down. Either the import carries on without
# error, or it aborts cleanly.
#
# The requirement is a disjunction, so this case asserts the disjunction. For
# every sheet, emptied in turn:
#
#   1. the outcome is one of the two documented shapes, "OK" or
#      "ERROR <number>: <description>" with a description that says something
#   2. a setup file is written, so the user has something rather than nothing
#   3. the sheets that were NOT emptied still land their data
#   4. Excel does not survive the operation
#
# Those four are what "did not break" means on disk. The failure modes they
# guard: a hang with no outcome at all, an outcome with an empty description, a
# run that writes nothing, an empty sheet costing the other four their data,
# and an Excel left running to be found by the next case.
#
# THE EXIT CODE OF THE DRIVER IS NOT THE VERDICT
# -----------------------------------------------------------------------------
# convert-setup.ps1 exits 1 when the converted setup's dictionary is empty,
# which is its own guard against a conversion that answers OK over nothing. So
# the Dictionary case and scope A below both exit 1 with an outcome of OK. This
# case reads the OUTCOME LINE, which is the product's answer; the driver's exit
# code is its own opinion about usefulness.
#
# WHAT WAS MEASURED, 2026-11-25, ebola-2026-11-19.xlsb into setup_dev.xlsm
# -----------------------------------------------------------------------------
# Every one of the five, emptied in turn, answered OK. None of them needed the
# "abort cleanly" half of the requirement, and in each the dictionary still
# landed its 381 rows:
#
#   sheet         cells cleared   outcome   dictionary rows after
#   Dictionary          4582        OK            381   (blank, see below)
#   Choices              781        OK            381
#   Exports               33        OK            381
#   Analysis             312        OK            381   (across 9 tables)
#   Translations        1084        OK            381
#
# The 381 rows on the Dictionary line are worth knowing: ClearContents empties
# the cells and leaves the table its height, so the import copied 381 EMPTY
# rows and answered OK. That is a faithful fixture of a user who selected their
# rows and pressed Delete. A user who deleted the table's ROWS would leave a
# table one row tall, which is a different shape and not what this builds.
#
# ALL FIVE SHEETS ARE PROTECTED, which the fixture has to get past: Dictionary
# and Choices leave their table bodies unlocked, Exports yields to Unprotect()
# with no password, and Analysis and Translations want the password the setup
# keeps in its own __pass sheet. The fixture reads that password off its own
# throwaway copy rather than carrying one, because a setup ships whatever
# password it ships.
#
# WHAT THIS CASE DOES NOT CLAIM
# -----------------------------------------------------------------------------
# That the result is USEFUL. A target carrying 381 blank dictionary rows
# answers every assertion here and is junk a user would have to clean up. The
# requirement is about the import surviving, and judging the usefulness of a
# conversion from an empty sheet is a separate question nobody has asked yet.
#
# It also does not claim the import is resilient sheet BY sheet. See case 040:
# the import raises on the first sheet that fails and does not collect a fault
# per sheet. This case empties ONE sheet at a time, so it never exercises two
# faults at once.

set -uo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/../lib/verify-lib.sh"

CASE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
FIXTURE_PS1="$CASE_DIR/050-empty-import-sheet.ps1"

# The conversion driver, in either home. Nothing in scripts/ drives
# HeadlessBuild.ImportSetupFromWorkbook yet; the driver lives in the headless
# draft folder until it is promoted into scripts/headless/windows/.
CONVERT="$(first_existing \
    "$REPO/scripts/headless/windows/convert-setup.ps1" \
    "$FIXTURES/convert-setup.ps1" || true)"

need_file "$CONVERT" "the conversion driver"
need_file "$FIXTURE_PS1" "the fixture builder"
need_file "$EMPTY_SETUP" "the current empty setup"

# A CLEAN old setup: one that converts whole, so an emptied sheet is the only
# thing wrong with it. malaria-aweil is deliberately not used here -- it has a
# fault of its own on Translations and case 040 owns that.
SOURCE="$(first_existing \
    "$FIXTURES/setups/old/ebola-2026-11-19.xlsb" \
    "$FIXTURES/setups/old/cholera-nga-borno.xlsb" \
    "$FIXTURES/setups/old/measles-training-2024.xlsb" || true)"
need_file "$SOURCE" "a clean old setup to empty a sheet of"

OUT="$VERIFY_WORK/050-empty-sheet"
mkdir -p "$OUT"

win() { cygpath -w "$1" 2>/dev/null || printf '%s' "$1"; }

# convert <source> <target> <log>: run the driver, answer nothing, leave the
# log for the readers below.
convert() {
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(win "$CONVERT")" \
        -Source "$(win "$1")" -Target "$(win "$2")" -Template "$(win "$EMPTY_SETUP")" \
        > "$3" 2>&1
    return 0
}

outcome_of() { grep -E '^(OK|ERROR )' "$1" 2>/dev/null | tr -d '\r' | tail -1; }
rows_of()    { sed -n 's/^rows=//p' "$1" 2>/dev/null | tr -d '\r' | head -1; }

# THE OUTCOME IS ONE OF THE TWO DOCUMENTED SHAPES. An "ERROR" with nothing
# after the colon is the failure this catches: it reports that something went
# wrong and says nothing a reader can act on.
assert_wellformed() {
    local outcome="$1"; local what="$2"
    if printf '%s' "$outcome" | grep -qE '^OK$'; then
        assert 0 "$what answered OK"
    elif printf '%s' "$outcome" | grep -qE '^ERROR -?[0-9]+: .+'; then
        assert 0 "$what answered a reported error: $outcome"
    else
        assert 1 "$what answered neither OK nor 'ERROR <n>: <text>' (got '${outcome:-nothing at all}')"
    fi
}

# =============================================================================
# SCOPE A -- every imported sheet empty at once
# =============================================================================
# The current empty setup imported into a fresh copy of itself. All five sheets
# are empty by construction, so this needs no fixture and runs on any machine
# that has the empty setup. It is the boundary: nothing to import at all.
step "scope A: the empty setup imported into a copy of itself"
A_TARGET="$OUT/scopeA-empty-into-empty.${EMPTY_SETUP##*.}"
A_LOG="$OUT/scopeA.log"
convert "$EMPTY_SETUP" "$A_TARGET" "$A_LOG"

A_OUTCOME="$(outcome_of "$A_LOG")"
say "outcome: ${A_OUTCOME:-nothing}   rows: $(rows_of "$A_LOG")"
assert_wellformed "$A_OUTCOME" "importing a wholly empty setup"
assert_file "$A_TARGET" "a setup was written even with nothing to import"

# =============================================================================
# SCOPE B -- one sheet empty, the other four full
# =============================================================================
# The valuable half: does one empty sheet cost the others their data?
SHEETS="Dictionary Choices Exports Analysis Translations"
COVERED=0
UNCOVERED=""

for sheet in $SHEETS; do
    say ""
    step "scope B: '$sheet' emptied, the other four left full"

    fixture="$OUT/src-no-$sheet.${SOURCE##*.}"
    target="$OUT/conv-no-$sheet.${EMPTY_SETUP##*.}"
    flog="$OUT/fixture-$sheet.log"
    clog="$OUT/convert-$sheet.log"

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(win "$FIXTURE_PS1")" \
        -Source "$(win "$SOURCE")" -Target "$(win "$fixture")" -Sheet "$sheet" \
        > "$flog" 2>&1
    fcode=$?

    if [ "$fcode" -ne 0 ]; then
        # A fixture that cannot be made is a gap in THIS case rather than a
        # fault in the product, so it is named loudly and not asserted on.
        # Every one of the five sheets is protected; Dictionary and Choices
        # leave their bodies unlocked, Exports yields to a bare Unprotect, and
        # Analysis and Translations need the password the setup keeps in its
        # own __pass sheet.
        say "NOT COVERED: the fixture could not empty '$sheet'"
        say "   $(grep -m1 '^ERROR' "$flog" | tr -d '\r')"
        UNCOVERED="$UNCOVERED $sheet"
        continue
    fi

    COVERED=$((COVERED + 1))
    say "fixture: $(grep -m1 '^CELLS_BEFORE=' "$flog" | tr -d '\r'), $(grep -m1 '^CELLS_AFTER=' "$flog" | tr -d '\r')"
    assert_contains "$flog" "EMPTIED=$sheet" "the fixture emptied '$sheet'"

    convert "$fixture" "$target" "$clog"
    outcome="$(outcome_of "$clog")"
    rows="$(rows_of "$clog")"
    say "outcome: ${outcome:-nothing}   dictionary rows: ${rows:-?}"

    assert_wellformed "$outcome" "an empty '$sheet'"
    assert_file "$target" "a setup was written with '$sheet' empty"

    # THE OTHER SHEETS SURVIVED. Asserted through the dictionary, which is the
    # one sheet whose row count the driver reports. Skipped when the dictionary
    # is the sheet that was emptied, because then there is nothing to survive.
    if [ "$sheet" != "Dictionary" ]; then
        assert_at_least "${rows:-0}" 100 \
            "an empty '$sheet' did not cost the dictionary its rows"
    fi
done

say ""
say "scope B covered $COVERED of 5 sheet(s)"
if [ -n "$UNCOVERED" ]; then
    say "NOT COVERED:$UNCOVERED -- read the NOT COVERED lines above"
fi

# A case that could empty nothing proves nothing, and saying so is better than
# a green that rests on scope A alone.
if [ "$COVERED" -eq 0 ]; then
    skip_case "the fixture could not empty any of the five sheets on this machine"
fi

# =============================================================================
# AND NOTHING WAS LEFT RUNNING
# =============================================================================
# An operation that reported a fault and left Excel behind did not abort
# cleanly. Inside a verify run every Excel is the suite's own, because the
# watchdog refuses to start while another is open.
left="$(wait_for_no_excel 20)"
assert_equal "$left" "0" "no Excel was left running after the conversions"

check_case "an empty imported worksheet is reported and does not break the import"

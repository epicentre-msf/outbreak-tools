#!/usr/bin/env bash
# title: Importing a structurally damaged setup lands what it can
# asserts: a clean old setup converts whole; a setup whose tables lost their rows, or a whole table, still lands its dictionary, writes a setup and builds
#
# WHAT THIS CASE IS ABOUT
# -----------------------------------------------------------------------------
# A user arrives with a setup they filled in months ago. Converting it means
# filling a fresh copy of the CURRENT empty setup from it, and an old file can
# carry anything: a sheet whose table moved, a table somebody deleted rows out
# of, a formula nobody fixed. The requirement is that such a setup does not
# take the whole import down with it -- what can land, lands, and what cannot
# is reported.
#
# WHY THE FAULT IS MANUFACTURED HERE
# -----------------------------------------------------------------------------
# Until 2026-11-25 this case converted malaria-aweil.xlsb, which raised error 91
# on its Translations sheet because Tab_Translations had no data body. Then that
# was FIXED -- case 060 covers it and passes -- and this case went red for the
# best possible reason: the bug it pinned was gone.
#
# A case that waits for a setup to be broken dies the day somebody fixes it, and
# takes its coverage with it quietly. So 040-delete-rows.ps1 makes the damage on
# a throwaway copy, in two shapes, and this case keeps asking the question after
# the next fix too.
#
# WHAT THE IMPORT ACTUALLY DOES, MEASURED 2026-11-25
# -----------------------------------------------------------------------------
# Over ebola-2026-11-19.xlsb (381 dictionary rows, 542 translations):
#
#   body rows deleted from Tab_TimeSeries_Analysis and Tab_Translations
#       -> OK, 381 rows. The five Analysis tables below TimeSeries had moved up
#          17 rows and the import found them anyway.
#   Tab_Translations deleted HEADER AND ALL, so no such table exists
#       -> OK, 381 rows.
#
# So the import absorbs both shapes. Everything that could land, landed, and the
# user gets a working setup out of a damaged one. That half of the requirement
# holds, and these assertions pin it.
#
# THE GAP, written down so nobody reads this pass as the requirement being met
# -----------------------------------------------------------------------------
# THE IMPORT REPORTS NOTHING. A source whose Tab_Translations is gone converts
# with outcome OK and not a word about the table that vanished, so the user gets
# a setup with no translations and no reason to suspect it. The header of this
# case used to call that out as the thing it guarded against -- "a silent
# partial import ... would hand back a setup missing a sheet and say nothing" --
# and that is now the measured behaviour.
#
# This case deliberately does NOT assert the silence. Asserting it would pin the
# bug, and pinning a bug is what killed the previous version of this case. The
# assertions take either branch -- a reported error or a clean absorb -- so they
# survive the fix, and the run PRINTS which branch happened. When the import
# learns to name a sheet it could not take, this case keeps passing and the
# printed line changes.
#
# Until then, "and what cannot is reported" is UNVERIFIED by this suite, because
# the product does not do it. Collecting a fault per sheet and naming it is the
# change the requirement asks for.

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

FIXTURE_PS1="$(dirname "${BASH_SOURCE[0]}")/040-delete-rows.ps1"
need_file "$FIXTURE_PS1" "the fixture builder"
need_file "$EMPTY_SETUP" "the current empty setup"

OLD_DIR="$FIXTURES/setups/old"
[ -d "$OLD_DIR" ] || skip_case "no old setups under $OLD_DIR"

win() { cygpath -w "$1" 2>/dev/null || printf '%s' "$1"; }

OUT="$VERIFY_WORK/040-import"
mkdir -p "$OUT"

# convert_one <source> <stem> -> prints "<exit>|<rows>|<outcome>|<target>"
convert_one() {
    local src="$1"
    local stem="$2"
    local target="$OUT/$stem.${EMPTY_SETUP##*.}"
    local log="$OUT/convert-$stem.log"

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(win "$CONVERT")" \
        -Source "$(win "$src")" -Target "$(win "$target")" \
        -Template "$(win "$EMPTY_SETUP")" \
        > "$log" 2>&1
    local code=$?
    local rows outcome
    rows="$(sed -n 's/^rows=//p' "$log" | head -1)"
    outcome="$(grep -E '^(OK|ERROR )' "$log" | tail -1)"
    printf '%s|%s|%s|%s\n' "$code" "${rows:-?}" "${outcome:-none}" "$target"
}

# make_fixture <source> <target> <tables> [--header] -> 0 when the damage landed
# The log is left beside the fixture so a case can quote it.
make_fixture() {
    local src="$1"; local target="$2"; local tables="$3"; local header="${4:-}"
    local log="$OUT/fixture-$(basename "${target%.*}").log"
    local extra=()
    [ "$header" = "--header" ] && extra+=(-IncludeHeader)

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(win "$FIXTURE_PS1")" \
        -Source "$(win "$src")" -Target "$(win "$target")" \
        -Tables "$tables" "${extra[@]}" \
        > "$log" 2>&1
    local code=$?
    sed 's/^/        /' "$log"
    return $code
}

# answered_cleanly <outcome>: OK, or an ERROR it chose to report. Anything else
# -- a crash, a dialog, a driver that never reached the import -- is the
# failure this case exists to catch.
answered_cleanly() {
    case "$1" in
        OK|"ERROR "*) return 0 ;;
        *) return 1 ;;
    esac
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

[ -n "$CLEAN_SRC" ] || skip_case "no clean old setup on this machine to damage a copy of"
say "the source every fixture is cut from: $(basename "$CLEAN_SRC")"

step "converting $(basename "$CLEAN_SRC") untouched, expected to convert whole"
IFS='|' read -r code rows outcome target <<< "$(convert_one "$CLEAN_SRC" "clean")"
say "exit $code, rows $rows, outcome: $outcome"
assert_equal "$code" "0" "a clean old setup converts and the driver exits 0"
assert_contains "$OUT/convert-clean.log" "^OK" "the import answered OK"
assert_at_least "$rows" 50 "the dictionary came across (>=50 rows)"
assert_file "$target" "the converted setup was written"
CLEAN_ROWS="$rows"

# --- 2. ROWS DELETED FROM TABLES, THE SHEET BELOW MOVED ---------------------
# Two tables on two sheets. TimeSeries is chosen because five more Analysis
# tables sit below it, so deleting its rows moves all five: that is the "a
# sheet whose table moved" fault in the case title.
step "fixture: body rows deleted from Tab_TimeSeries_Analysis and Tab_Translations"
MOVED_SRC="$OUT/src-rows-deleted.${CLEAN_SRC##*.}"
if make_fixture "$CLEAN_SRC" "$MOVED_SRC" "Analysis:Tab_TimeSeries_Analysis,Translations:Tab_Translations"; then
    assert_contains "$OUT/fixture-src-rows-deleted.log" "AFTER=Nothing" "the fixture emptied the tables it was asked to"

    step "converting the setup whose tables lost their rows"
    IFS='|' read -r mcode mrows moutcome mtarget <<< "$(convert_one "$MOVED_SRC" "rows-deleted")"
    say "exit $mcode, rows $mrows, outcome: $moutcome"

    answered_cleanly "$moutcome"
    assert $? "the import answered OK or a reported error rather than crashing"
    assert_at_least "$mrows" "$CLEAN_ROWS" "the whole dictionary still landed ($CLEAN_ROWS rows)"
    assert_file "$mtarget" "a setup was written from the damaged source"

    if [ "$moutcome" = "OK" ]; then
        say "the import reported nothing about the two emptied tables"
    else
        say "the import reported: $moutcome"
    fi

    # --- AND THE RESULT IS STILL USABLE ------------------------------------
    # The point of landing what it can: a linelist can be built from it. This
    # is the assertion that makes "it did not break the import" mean something
    # to a user rather than only to a log reader.
    if [ -f "$mtarget" ] && [ -f "$DESIGNER" ] && [ -f "$TEMPLATE" ]; then
        step "building a linelist from the damaged source's setup"
        build_linelist "$OUT" "ll-damaged" "$mtarget" --lllang=ENG-English
        say "build exited $?"
        assert_file "$OUT/ll-damaged.xlsm" "a linelist builds from the setup imported out of a damaged source"
        if [ -f "$OUT/ll-damaged-generation.txt" ]; then
            # grep -c already prints a count and exits 1 when that count is
            # zero, so an `|| echo 0` here appends a SECOND zero and the say
            # below breaks across two lines. Swallow the status instead.
            n="$(grep -c "\[Error\]" "$OUT/ll-damaged-generation.txt" 2>/dev/null)" || true
            say "the generation logged ${n:-0} error(s), which belong to the source's own formulas"
        fi
    else
        say "no designer or template, so the usability half is skipped"
    fi
else
    say "the fixture could not delete those rows on this machine; this scope is skipped"
fi

# --- 3. A TABLE GONE ENTIRELY, HEADER AND ALL -------------------------------
# The harder shape: no Tab_Translations on the sheet at all. An import looking
# for it by name finds nothing.
step "fixture: Tab_Translations deleted header and all"
GONE_SRC="$OUT/src-table-gone.${CLEAN_SRC##*.}"
if make_fixture "$CLEAN_SRC" "$GONE_SRC" "Translations:Tab_Translations" --header; then
    assert_contains "$OUT/fixture-src-table-gone.log" "STILL_THERE=no" "the fixture removed the table outright"

    step "converting the setup whose Translations table is gone"
    IFS='|' read -r gcode grows goutcome gtarget <<< "$(convert_one "$GONE_SRC" "table-gone")"
    say "exit $gcode, rows $grows, outcome: $goutcome"

    answered_cleanly "$goutcome"
    assert $? "a missing table does not crash the import"
    assert_at_least "$grows" "$CLEAN_ROWS" "the dictionary landed whole despite the missing table"
    assert_file "$gtarget" "a setup was written although a table was gone"

    if [ "$goutcome" = "OK" ]; then
        say ""
        say "SILENT: Tab_Translations does not exist in the source and the import"
        say "answered OK without naming it. The user gets a setup with no"
        say "translations and nothing to tell them why. See THE GAP in this"
        say "case's header; the assertions above take either branch on purpose."
        say ""
    else
        say "the import named the missing table: $goutcome"
    fi
else
    say "the fixture could not remove that table on this machine; this scope is skipped"
fi

# --- 4. THE EMPTY SETUP, WITH ROWS DELETED FROM SOME TABLES -----------------
# A source with nothing in it to begin with, damaged on top. Judged on the
# OUTCOME rather than the exit code: convert-setup.ps1 exits 1 whenever the
# dictionary lands zero rows, and an empty source always lands zero, so the
# exit code says nothing about whether the import coped.
step "fixture: the empty setup, rows deleted from four tables"
EMPTY_SRC="$OUT/src-empty-rows-deleted.${EMPTY_SETUP##*.}"
if make_fixture "$EMPTY_SETUP" "$EMPTY_SRC" \
        "Translations:Tab_Translations,Choices:Tab_Choices,Analysis:Tab_Global_Summary,Analysis:Tab_TimeSeries_Analysis"; then
    step "converting the damaged empty setup"
    IFS='|' read -r ecode erows eoutcome etarget <<< "$(convert_one "$EMPTY_SRC" "empty-rows-deleted")"
    say "exit $ecode, rows $erows, outcome: $eoutcome"
    say "the driver exits 1 on zero dictionary rows by its own rule, so the outcome is what counts here"

    answered_cleanly "$eoutcome"
    assert $? "a damaged empty setup answered OK or a reported error"
    assert_file "$etarget" "a setup was written even from a damaged empty source"
else
    say "the fixture could not delete those rows on this machine; this scope is skipped"
fi

left="$(wait_for_no_excel 20)"
assert_equal "$left" "0" "no Excel was left running after the conversions"

check_case "a damaged setup lands what it can, and the import is silent about what it drops"

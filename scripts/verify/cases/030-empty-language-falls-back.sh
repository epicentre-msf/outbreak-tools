#!/usr/bin/env bash
# title: A setup language whose translation column is empty falls back to a filled one
# asserts: the linelist built on the empty column carries the same labels as the baseline
#
# WHAT THIS CASE IS ABOUT
# -----------------------------------------------------------------------------
# A setup's translation table has one column per language, named by the
# language: English, Portuguese, Français. RNG_LangSetup names the column the
# dictionary, the choices and the analyses are written from. A user can pick a
# language whose column EXISTS AND IS EMPTY -- a column added for a translation
# nobody finished.
#
# The required behaviour: the build falls back to a filled column, so the
# delivered linelist reads in a language. The failure it guards against is a
# linelist whose every label is blank, which the build reports OK over: the
# lookup returns an empty string for each tag and nothing says so.
#
# This is NOT the lllang fault. lllang is the INTERFACE language, a CODE
# (ENG, FRA, SPA, POR, ARA) keyed on the designer's own translation tables, and
# linelist-inspect.R already catches a code that names no column -- worksheets
# come out called LLSHEET_* and RNG_LLLanguageCode holds a name where a code
# belongs. setuplang is the setup's OWN column name and a different vocabulary.
# scripts/headless/README.md has the long version. The two are confused easily
# and silently, which is why this case says which one it tests.
#
# HOW IT JUDGES, WITHOUT OPENING EXCEL
# -----------------------------------------------------------------------------
# Two builds from the same setup:
#   baseline  the untouched setup, built on a FILLED column
#   test      a copy with one column emptied, built on THAT column
#
# A workbook is a zip, and its labels are in xl/sharedStrings.xml. If the
# fallback worked, the strings the baseline carries are in the test build too.
# If it did not, they are simply absent, because an empty label is written as
# nothing at all. So the overlap of the two string tables is the measurement,
# and it needs no Excel and no judgment.

set -uo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/../lib/verify-lib.sh"

need_command Rscript
need_command unzip
need_file "$DESIGNER" "the designer mock"
need_file "$TEMPLATE" "the ribbon template"

# A CURRENT setup with clean content is wanted here. measles carries no
# deliberately broken formula, so its string table is all real labels; the
# generic fixture is the fallback because it is the one a clone has.
SETUP=""
for candidate in "$FIXTURES/setups/current/measles.xlsb" \
                 "$FIXTURES/setups/current/cholera.xlsb" \
                 "$GENERIC_SETUP"
do
    if [ -n "$candidate" ] && [ -f "$candidate" ]; then SETUP="$candidate"; break; fi
done
need_file "$SETUP" "a current setup to build from"

OUT="$VERIFY_WORK/030-language"
mkdir -p "$OUT"

# The column to empty, and the one the baseline is built on. Both are read off
# the setup rather than assumed: a setup ships the languages it ships.
BLANK_LANG="${OBT_VERIFY_BLANK_LANG:-}"
BASE_LANG="${OBT_VERIFY_BASE_LANG:-English}"

# 1. THE FIXTURE. The .ps1 reports the table, the columns and the counts, so a
#    wrong guess at a language name fails here with the real list.
FIXTURE="$OUT/setup-blank-language.${SETUP##*.}"
if [ -z "$BLANK_LANG" ]; then
    # Français is in every setup this project ships and is never the default,
    # so emptying it leaves a filled column to fall back to.
    BLANK_LANG="Français"
fi

step "blanking the '$BLANK_LANG' column of $(basename "$SETUP")"
powershell.exe -NoProfile -ExecutionPolicy Bypass \
    -File "$(dirname "${BASH_SOURCE[0]}")/030-blank-language-column.ps1" \
    -Source "$(cygpath -w "$SETUP" 2>/dev/null || echo "$SETUP")" \
    -Target "$(cygpath -w "$FIXTURE" 2>/dev/null || echo "$FIXTURE")" \
    -Language "$BLANK_LANG" \
    > "$OUT/fixture.log" 2>&1
FIXTURE_EXIT=$?
sed 's/^/      /' "$OUT/fixture.log"

if [ "$FIXTURE_EXIT" -ne 0 ]; then
    skip_case "the fixture could not be made: $(grep -m1 '^ERROR' "$OUT/fixture.log" | head -1)"
fi
assert_file "$FIXTURE" "the fixture setup was written"

# ASSERT ON THE COUNT, NEVER ON THE LANGUAGE NAME. A language name holds
# accents, the fixture's stdout crosses into bash through the console code page,
# and grepping for "BLANKED=Français" against a log that spelled the cedilla in
# another encoding is how this case failed on 2026-11-25 while the column had in
# fact been emptied. BLANKED_FILLED_AFTER is ASCII and is the number the
# question turns on.
FILLED_AFTER="$(sed -n 's/^BLANKED_FILLED_AFTER=//p' "$OUT/fixture.log" | head -1)"
FILLED_BEFORE="$(sed -n 's/^BLANKED_FILLED_BEFORE=//p' "$OUT/fixture.log" | head -1)"
assert_equal "${FILLED_AFTER:-unset}" "0" "the '$BLANK_LANG' column was emptied"
assert_at_least "${FILLED_BEFORE:-0}" 1 "the column had something in it to empty"

ROWS="$(sed -n 's/^ROWS=//p' "$OUT/fixture.log" | head -1)"
say "the table held ${ROWS:-?} translation row(s)"

# 2. THE BASELINE, from the untouched setup on a filled column.
build_linelist "$OUT" "ll-baseline" "$SETUP" \
    "--setuplang=$BASE_LANG" --lllang=ENG-English
say "baseline build exit $?"

# 3. THE TEST, from the fixture on the emptied column.
build_linelist "$OUT" "ll-blank" "$FIXTURE" \
    "--setuplang=$BLANK_LANG" --lllang=ENG-English
say "test build exit $?"

BASELINE="$OUT/ll-baseline.xlsm"
TEST="$OUT/ll-blank.xlsm"

assert_file "$BASELINE" "the baseline linelist was built"
assert_file "$TEST" "the linelist on the empty column was built at all"

if [ ! -f "$BASELINE" ] || [ ! -f "$TEST" ]; then
    check_case
fi

# 4. THE MEASUREMENT. Label-like strings of the baseline, and how many of them
#    the test build also carries.
strings_of "$BASELINE" | sort -u > "$OUT/baseline-strings.txt"
strings_of "$TEST"     | sort -u > "$OUT/test-strings.txt"

# A label is a string with letters and some length. Numbers, single characters
# and formula fragments are dropped: they would be shared by both builds
# whatever happened to the labels and would flatter the overlap.
grep -E '[A-Za-zÀ-ÿ]{4}' "$OUT/baseline-strings.txt" \
    | grep -vE '^[0-9.,-]+$' | sort -u > "$OUT/baseline-labels.txt"

BASE_N=$(wc -l < "$OUT/baseline-labels.txt" | tr -d '[:space:]')
TEST_N=$(wc -l < "$OUT/test-strings.txt" | tr -d '[:space:]')
KEPT=$(comm -12 "$OUT/baseline-labels.txt" <(sort -u "$OUT/test-strings.txt") | wc -l | tr -d '[:space:]')

say "baseline label-like strings: $BASE_N"
say "test build total strings:    $TEST_N"
say "of the baseline's labels, the test build carries: $KEPT"

assert_at_least "$BASE_N" 20 "the baseline carries a real label set to compare against"

if [ "${BASE_N:-0}" -gt 0 ]; then
    PCT=$(( KEPT * 100 / BASE_N ))
    say "overlap: ${PCT}%"
    # A working fallback reproduces the filled column's labels, so the overlap
    # is high. A fallback that did not happen leaves the labels empty and the
    # overlap near zero. 50% separates the two without pinning the exact
    # column the fallback chooses.
    assert_at_least "$PCT" 50 "the labels of a filled column came through (>=50% overlap)"
    if [ "$PCT" -lt 50 ]; then
        say "the labels the test build is MISSING, first 10:"
        comm -23 "$OUT/baseline-labels.txt" <(sort -u "$OUT/test-strings.txt") \
            | head -10 | sed 's/^/      /'
        say "this is the empty-label fault: the build answered OK over a linelist with no labels"
    fi
fi

# 5. AND IT STILL HAS TO BE A COHERENT LINELIST. A fallback that supplied
#    labels while breaking something else is no pass.
inspect_linelist "$TEST" "$OUT/ll-blank-inspect.log"
assert $? "linelist-inspect.R reports the fallback build clean"

check_case "an empty translation column falls back to a filled one"

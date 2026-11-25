#!/usr/bin/env bash
# title: The generic setup builds headless, and what comes out is coherent
# asserts: a linelist is written, linelist-inspect is clean, only the brok_* fixture errors appear
#
# WHY THIS CASE
# -----------------------------------------------------------------------------
# generic-test-setup.xlsb is the setup every other part of this project leans
# on: TestHeadlessSetupFill fills from it, TestHeadlessLinelistBuild builds
# from what it filled, and the obt package ships it as inst/extdata. If a
# linelist cannot be built from it, nothing downstream is worth running.
#
# It is also the setup that carries DictionaryTestFixture's brok_* variables,
# labelled "This formula should fail". So its generation log ALWAYS holds
# [Error] lines, the launcher always calls the delivery untrustworthy, and
# build-linelist.R always exits 1. That is correct behaviour on this setup and
# the case allows exactly those errors -- no more, which is the point: an
# allowance of "any error" would pass a build that had broken everything else.
#
# WHAT COHERENT MEANS HERE
# -----------------------------------------------------------------------------
# The file on disk, and then scripts/devtools/linelist-inspect.R over it. That
# script is the repository's own judge of a delivered linelist and it reads the
# bytes: the three classes that must travel, no FormLogic module, no
# placeholder worksheet, no worksheet still called LLSHEET_*, an
# RNG_LLLanguageCode that names a real translation column, ribbon callbacks
# that exist, and the twelve geo dictionary rows. Keeping one definition of
# "coherent" in the repository is worth more than a second one here.

set -uo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/../lib/verify-lib.sh"

need_command Rscript
need_file "$GENERIC_SETUP" "the generic setup"
need_file "$DESIGNER" "the designer mock"
need_file "$TEMPLATE" "the ribbon template"

OUT="$VERIFY_WORK/010-generic"
NAME="ll-verify-generic"

# No --geopath: the generic setup carries no geography, and an empty geobase is
# valid and common. The geo assertions of linelist-inspect are about a setup
# that ASKED for geography, so a setup with none passes them by.
build_linelist "$OUT" "$NAME" "$GENERIC_SETUP" \
    --setuplang=English --lllang=ENG-English
BUILD_EXIT=$?

LINELIST="$OUT/$NAME.xlsm"
REPORT="$OUT/$NAME-build-report.txt"
GENLOG="$OUT/$NAME-generation.txt"

# 1. A FILE, and one with a workbook's worth of bytes in it.
assert_file "$LINELIST" "a linelist was written as .xlsm"
assert_file_bigger_than "$LINELIST" 1000000 "the linelist is a real workbook"

# 2. WHAT THE BUILD ITSELF RECORDED. outcome=OK means the build ran to the
#    end. It says nothing about the size of what was built, which is why the
#    sheet and variable counts are read too.
assert_file "$REPORT" "the build wrote its report"
if [ -f "$REPORT" ]; then
    assert_equal "$(report_value "$REPORT" outcome)" "OK" "the build answered OK"
    assert_at_least "$(report_value "$REPORT" sheets)" 1 "at least one data entry sheet was built"
    assert_at_least "$(report_value "$REPORT" variables)" 50 "the dictionary came across"
    assert_at_least "$(report_value "$REPORT" components)" 70 "the source components were re-imported"
    say "report: $(report_value "$REPORT" sheets) sheet(s), $(report_value "$REPORT" variables) variable(s), $(report_value "$REPORT" components) component(s)"
fi

# 3. THE ERRORS, AND ONLY THE EXPECTED ONES.
if [ -f "$GENLOG" ]; then
    total=$(grep -c "\[Error\]" "$GENLOG" 2>/dev/null || echo 0)
    unexpected=$(grep "\[Error\]" "$GENLOG" 2>/dev/null | grep -vcE "$BROK_FIXTURE_ERRORS" || true)
    unexpected=${unexpected:-0}
    say "generation log: $total error(s), $unexpected of them unexpected"
    assert_equal "$unexpected" "0" "every generation error is a known brok_* fixture"
    if [ "$unexpected" != "0" ]; then
        say "the unexpected ones:"
        grep "\[Error\]" "$GENLOG" | grep -vE "$BROK_FIXTURE_ERRORS" | head -5 | sed 's/^/      /'
    fi
fi

# 4. THE REPOSITORY'S OWN JUDGE of a delivered linelist.
if [ -f "$LINELIST" ]; then
    inspect_linelist "$LINELIST" "$OUT/$NAME-inspect.log"
    assert $? "linelist-inspect.R reports the delivered file clean"
    if [ -s "${VERIFY_LAST_INSPECT_LOG:-}" ]; then
        grep -iE "fail|missing|absent" "${VERIFY_LAST_INSPECT_LOG}" | head -6 | sed 's/^/      /'
    fi
fi

# BUILD_EXIT is deliberately NOT asserted. build-linelist.R exits 1 on this
# setup by design: its F5 guard refuses to call a delivery trustworthy when the
# generation recorded errors, and the brok_* fixtures guarantee some. The
# assertions above are what judge the build instead.
say "build-linelist.R exit was $BUILD_EXIT (expected 1 on this setup: the brok_* fixtures)"

check_case "the generic setup builds and the delivered linelist is coherent"

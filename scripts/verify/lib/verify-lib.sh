#!/usr/bin/env bash
# verify-lib.sh -- what every verify case is given.
#
# Sourced by verify.sh before a case runs, and by a case run on its own. It
# carries the paths, the assertions and the two wrappers around a build, so a
# case file holds its own argument and nothing else.
#
# A CASE IS A BEHAVIOUR, NOT A UNIT
# -----------------------------------------------------------------------------
# These cases generate real files with real Excel and judge what came out.
# They do not test a function in isolation: src/tests/ does that, and
# `Rscript scripts/tests/run-tests.R` runs it. A case here asks whether the
# PRODUCT is coherent -- whether a linelist a person would open reads in a
# language, carries its geography, compiles and holds the sheets its setup
# asked for.
#
# THE CONTRACT A CASE KEEPS
# -----------------------------------------------------------------------------
#   exit 0   PASS  every assertion held
#   exit 2   SKIP  a precondition is absent, and the case says which
#   other    FAIL  an assertion did not hold, and the case says which
#
# A case prints what it did as it goes. verify.sh keeps that output in a log
# per case and prints one verdict line each, so a reader never has to open a
# log to see what happened.
#
# WHY SKIP IS A FIRST-CLASS ANSWER
# -----------------------------------------------------------------------------
# Every fixture this suite needs is a binary workbook, and `.gitignore` keeps
# every one of them off git: `src/bin/`, `.mock/` and `src/tests/.input/**.xlsb`
# are all ignored. So a clone has the scripts and none of the inputs. A case
# that FAILED on a missing fixture would make the suite red on any machine but
# one, and a red that means "nothing to test here" teaches people to ignore
# reds. SKIP says it plainly and keeps the run green.

set -uo pipefail

# --- where everything is ------------------------------------------------------
VERIFY_LIB_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VERIFY_DIR="$(dirname "$VERIFY_LIB_DIR")"
REPO="$(cd "$VERIFY_DIR/../.." && pwd)"

# The fixture root. Every setup a case builds from lives under it, and
# OBT_VERIFY_FIXTURES moves the lot for a machine that keeps them elsewhere.
FIXTURES="${OBT_VERIFY_FIXTURES:-$REPO/.obt/draft/headless}"

# Where a case writes. verify.sh sets this per run; a case run on its own makes
# its own.
VERIFY_WORK="${VERIFY_WORK:-$REPO/.obt/draft/verify/adhoc-$(date +%Y%m%d-%H%M%S)}"

# first_existing <path>...: the first path on disk, in the order GIVEN.
#
# `ls a b | head -1` was the first try and it is wrong: ls SORTS its arguments,
# so `ls x.xlsm x.xlsb` answers the .xlsb and the suite tested the format the
# project moved away from. Preference order has to be the caller's.
first_existing() {
    local candidate
    for candidate in "$@"; do
        if [ -n "$candidate" ] && [ -f "$candidate" ]; then
            printf '%s' "$candidate"
            return 0
        fi
    done
    return 1
}

# The three inputs almost every build needs, each taking .xlsm first because
# that is what the project ships now.
DESIGNER="$(first_existing \
    "$REPO/.mock/designer_mock.xlsm" \
    "$REPO/.mock/xlsb/designer_mock.xlsb" \
    "$REPO/.mock/designer_mock.xlsb" || true)"

TEMPLATE="$(first_existing \
    "$REPO/ribbons/_ribbontemplate_dev.xlsm" \
    "$REPO/ribbons/_ribbontemplate_dev.xlsb" || true)"

EMPTY_SETUP="$(first_existing \
    "$REPO/src/bin/setup/setup_dev.xlsm" \
    "$REPO/src/bin/setup/setup_dev.xlsb" \
    "$FIXTURES/template/setup-empty-current.xlsm" \
    "$FIXTURES/template/setup-empty-current.xlsb" || true)"

# The generic setup, the one case 020 and case 010 both build from. The repo
# fixture is preferred over the harvested copy: it is the file the suites use.
GENERIC_SETUP="$(first_existing \
    "$REPO/src/tests/.input/package/generic-test-setup.xlsb" \
    "$FIXTURES/setups/old/generic.xlsb" || true)"

VERIFY_FAILURES=0

# --- saying things ------------------------------------------------------------
if [ -t 1 ] && [ -z "${NO_COLOR:-}" ]; then
    C_FAIL=$'\033[1;31m'; C_OK=$'\033[32m'; C_SKIP=$'\033[33m'; C_OFF=$'\033[0m'
else
    C_FAIL=""; C_OK=""; C_SKIP=""; C_OFF=""
fi

say()  { printf '    %s\n' "$*"; }
step() { printf '  > %s\n' "$*"; }

# A case ends through one of these three and nothing else.
pass_case() { printf '  %sPASS%s %s\n' "$C_OK" "$C_OFF" "${1:-}"; exit 0; }
skip_case() { printf '  %sSKIP%s %s\n' "$C_SKIP" "$C_OFF" "${1:-}"; exit 2; }
fail_case() { printf '  %sFAIL%s %s\n' "$C_FAIL" "$C_OFF" "${1:-}"; exit 1; }

# --- preconditions ------------------------------------------------------------
# need_file <path> <what it is>: SKIP, never FAIL. See the note at the top.
need_file() {
    if [ -z "${1:-}" ] || [ ! -f "$1" ]; then
        skip_case "no ${2:-fixture} on this machine${1:+ at $1}"
    fi
    say "using ${2:-fixture}: $(basename "$1")"
}

need_command() {
    command -v "$1" >/dev/null 2>&1 || skip_case "$1 is not on PATH"
}

# --- assertions ---------------------------------------------------------------
# Each records a failure and carries on, so one case reports every fault it
# found rather than the first. check_case at the end turns the count into the
# exit code.
assert() {
    local ok="$1"; local what="$2"
    if [ "$ok" -eq 0 ]; then
        printf '    ok   %s\n' "$what"
    else
        printf '    %sBAD%s  %s\n' "$C_FAIL" "$C_OFF" "$what"
        VERIFY_FAILURES=$((VERIFY_FAILURES + 1))
    fi
}

assert_file() {
    [ -f "$1" ]; assert $? "${2:-a file exists at $1}"
}

assert_file_bigger_than() {
    local path="$1"; local min="$2"; local what="${3:-$(basename "$path") is at least ${min} bytes}"
    local size=0
    [ -f "$path" ] && size=$(stat -c %s "$path" 2>/dev/null || echo 0)
    [ "$size" -ge "$min" ]; assert $? "$what (got $size)"
}

assert_contains() {
    grep -q -- "$2" "$1" 2>/dev/null; assert $? "${3:-$(basename "$1") contains '$2'}"
}

assert_not_contains() {
    ! grep -q -- "$2" "$1" 2>/dev/null; assert $? "${3:-$(basename "$1") is free of '$2'}"
}

assert_equal() {
    [ "$1" = "$2" ]; assert $? "${3:-'$1' equals '$2'}"
}

assert_at_least() {
    [ "${1:-0}" -ge "${2:-0}" ] 2>/dev/null; assert $? "${3:-$1 is at least $2}"
}

# check_case: the last line of every case.
check_case() {
    if [ "$VERIFY_FAILURES" -eq 0 ]; then
        pass_case "${1:-}"
    fi
    fail_case "$VERIFY_FAILURES assertion(s) did not hold"
}

# --- reading what a build recorded -------------------------------------------
report_value() {
    # report_value <build-report.txt> <key>
    sed -n "s/^$2=//p" "$1" 2>/dev/null | head -1
}

# The [Error] lines of a generation log that are NOT expected.
#
# Two setups in this project carry deliberately broken formulas:
# demo_setup_filled and generic-test-setup both come from
# DictionaryTestFixture.bas, whose brok_* variables are labelled "This formula
# should fail". Their logs always hold [Error] lines, and a case built on them
# has to allow exactly those and no others -- an allowance of "any error" would
# pass a build that broke everything.
unexpected_errors() {
    # unexpected_errors <generation.txt> [allowed-regex]
    local log="$1"; local allowed="${2:-}"
    [ -f "$log" ] || { echo 0; return; }
    if [ -n "$allowed" ]; then
        grep -c "\[Error\]" "$log" 2>/dev/null >/dev/null
        grep "\[Error\]" "$log" 2>/dev/null | grep -vcE "$allowed" || true
    else
        grep -c "\[Error\]" "$log" 2>/dev/null || true
    fi
}

# The brok_* fixtures, as one regex. Any case building from the generic or the
# demo setup passes this to unexpected_errors.
BROK_FIXTURE_ERRORS='brok_form_v1|brok_grp_tab_v1|brok_grp_arg_v1|brok_grp_agg_v1|Unknown token .choi_v2.|Grouped formulas? (require|must)|GROUP_NOPE'

# --- the two wrappers around a build -----------------------------------------
# build_linelist <out-dir> <name> <setup> [extra build-linelist.R args...]
#
# R_PROFILE_USER is set because the repository root carries an .Rprofile that
# loads pacman and the tidyverse, and Rscript sources it before the build's
# first line. Without this the build dies over packages it never asked for.
build_linelist() {
    local out="$1"; local name="$2"; local setup="$3"; shift 3
    mkdir -p "$out"
    step "building $name from $(basename "$setup")"
    R_PROFILE_USER=nonexistent Rscript "$REPO/scripts/headless/build-linelist.R" \
        "--setup=$setup" \
        "--out=$out" \
        "--name=$name" \
        "--designer=$DESIGNER" \
        "--temppath=$TEMPLATE" \
        "$@" \
        > "$out/$name-run.log" 2>&1
    local code=$?
    say "exit $code, log $out/$name-run.log"
    return $code
}

# inspect_linelist <linelist>: the repository's own judge of a delivered file.
#
# scripts/devtools/linelist-inspect.R reads the zip and reports the faults a
# generation answers OK over: the three classes that must travel, a FormLogic
# module that must not, the placeholder worksheet, worksheets still called
# LLSHEET_* and an RNG_LLLanguageCode naming no translation column, the ribbon
# callbacks, and the twelve geo dictionary rows. Using it here keeps one
# definition of "a coherent linelist" in the repository.
inspect_linelist() {
    local target="$1"
    local log="${2:-$VERIFY_WORK/inspect-$(basename "$target").log}"
    step "inspecting $(basename "$target")"
    R_PROFILE_USER=nonexistent Rscript "$REPO/scripts/devtools/linelist-inspect.R" \
        "$target" > "$log" 2>&1
    local code=$?
    say "inspect exit $code, log $log"
    VERIFY_LAST_INSPECT_LOG="$log"
    return $code
}

# --- Excel, after a driver says it is done -----------------------------------
# An operation that reported a fault and left Excel running did not abort
# cleanly: the next case then finds an instance it did not start, and on this
# laptop a forgotten Excel is how memory runs out. The watchdog refuses to
# start the suite while any Excel is open, so inside a run every instance is
# the suite's own and a count of zero is the right expectation.
excel_count() {
    tasklist //FI "IMAGENAME eq EXCEL.EXE" 2>/dev/null | grep -ci 'excel\.exe' || true
}

# wait_for_no_excel [seconds]: give Excel a moment to go, then answer the count.
# Excel takes a second or two to exit after a COM release, so a bare count
# right after a driver returns reads as an orphan that is merely slow.
wait_for_no_excel() {
    local limit="${1:-20}"
    local waited=0
    while [ "$waited" -lt "$limit" ]; do
        if [ "$(excel_count)" -eq 0 ]; then break; fi
        sleep 2
        waited=$((waited + 2))
    done
    excel_count
}

# --- the shared string table of a delivered workbook -------------------------
# A linelist written from an EMPTY translation column comes out with empty
# labels, and the strings that should have been there are simply absent from
# the file. So the string table is what tells one build's labels from
# another's, and it needs no Excel: a workbook is a zip.
#
# strings_of <workbook> -> one string per line, on stdout.
strings_of() {
    local target="$1"
    local tmp
    tmp="$(mktemp -d)"
    if unzip -o -q "$target" "xl/sharedStrings.xml" -d "$tmp" 2>/dev/null; then
        tr '<' '\n' < "$tmp/xl/sharedStrings.xml" \
            | sed -n 's/^t[^>]*>//p' \
            | sed 's/&amp;/\&/g; s/&lt;/</g; s/&gt;/>/g; s/&quot;/"/g; s/&apos;/'"'"'/g' \
            | sed '/^$/d'
    fi
    rm -rf "$tmp"
}

export -f strings_of 2>/dev/null || true

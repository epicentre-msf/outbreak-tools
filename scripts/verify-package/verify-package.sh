#!/usr/bin/env bash
# verify-package.sh [--list] [--setup <pattern>] [--records <n>] [--keep]
#                   [--no-guard] [--help]
#
# Drive every headless setup through the obt PACKAGE: generate its linelist,
# make a fake dataset out of the same setup, and import that dataset into the
# linelist. One line per setup, and the exit code is the verdict of the run.
#
# WHAT THIS ANSWERS, AND WHAT ANSWERS THE REST
# -----------------------------------------------------------------------------
#   scripts/tests/run-tests.R      the UNIT suites, inside Excel, over src/
#   scripts/release/preflight.sh   the STATIC release gate
#   scripts/verify/verify.sh       the product's BEHAVIOUR, end to end, over src/
#   scripts/verify-package/        THIS: the shipped PACKAGE over the shipped
#                                  BINARIES, with data put through the result
#
# verify.sh builds with the mock designer carrying src/, so it judges the code
# in the working tree. This judges what a user actually installs: the obt
# package driving a downloaded designer. Case 020 of verify.sh used to ask a
# thin version of this question and was removed, because it could only ever run
# against whatever channel happened to be cached and so was permanently stale.
#
# THE TWO GATES, AND WHY THEY SKIP RATHER THAN FAIL
# -----------------------------------------------------------------------------
# 1. THE PACKAGE must be the latest. A run over a stale obt reports on a package
#    its author has already moved past, and a red then looks like a product
#    fault when the fix is merely not installed here.
# 2. THE CHANNEL must be one obt can drive. obt accepts a dev channel stamped
#    "latest" and a semantic version at or above its minimum; a date-stamped zip
#    from the old scheme it aborts on.
#
# Either gate failing means the question was NOT ASKED. That is a SKIP, with the
# reason and the command that fixes it, the same way a missing fixture skips in
# verify.sh. A red that every machine shows is a red people learn to ignore.
#
# CLEANING UP AFTER ITSELF
# -----------------------------------------------------------------------------
# Each setup gets a working tree holding a copied channel, a setup, a generated
# linelist and a fake dataset: tens of megabytes apiece. A green run DELETES all
# of it and keeps the logs, because the next run rebuilds it and the logs are the
# record that it was green. A failed run keeps everything, because what a broken
# run left behind is the only account of what went wrong. --keep keeps both.

set -uo pipefail

PKG_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# PKG_DIR is <repo>/scripts/verify-package, so the root is TWO levels up.
REPO="$(cd "$PKG_DIR/../.." && pwd)"

# The assert vocabulary, inspect_linelist and the Excel helpers all come from
# the verify suite's lib. One definition of "ok" and "BAD" across both suites
# beats a second copy that drifts.
VERIFY_LIB="$REPO/scripts/verify/lib/verify-lib.sh"
if [ ! -f "$VERIFY_LIB" ]; then
    printf 'verify-package.sh: no %s\n' "$VERIFY_LIB" >&2
    exit 2
fi

LIST=0
PATTERN=""
KEEP=0
NO_GUARD=0
RECORDS="${OBT_PACKAGE_RECORDS:-25}"

while [ $# -gt 0 ]; do
    case "$1" in
        --list) LIST=1 ;;
        --setup) PATTERN="${2:-}"; shift ;;
        --records) RECORDS="${2:-25}"; shift ;;
        --keep) KEEP=1 ;;
        --no-guard) NO_GUARD=1 ;;
        --help|-h)
            sed -n '2,46p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'
            exit 0 ;;
        *) printf 'verify-package.sh: unknown option %s\n' "$1" >&2; exit 2 ;;
    esac
    shift
done

C_OK=$'\033[32m'; C_FAIL=$'\033[31m'; C_SKIP=$'\033[33m'; C_OFF=$'\033[0m'
if [ ! -t 1 ]; then C_OK=""; C_FAIL=""; C_SKIP=""; C_OFF=""; fi

FIXTURES="${OBT_VERIFY_FIXTURES:-$REPO/.obt/draft/headless}"

# --- the setups --------------------------------------------------------------
# The CURRENT setups are what a user has: the ones the headless folder tracks as
# convertible and up to date. OBT_PACKAGE_SETUPS points somewhere else.
SETUP_DIR="${OBT_PACKAGE_SETUPS:-$FIXTURES/setups/current}"
SETUPS=()
if [ -d "$SETUP_DIR" ]; then
    while IFS= read -r s; do SETUPS+=("$s"); done < <(
        find "$SETUP_DIR" -maxdepth 1 -type f \( -name '*.xlsb' -o -name '*.xlsm' \) | sort
    )
fi

if [ "$LIST" -eq 1 ]; then
    printf '\nThe setups verify-package drives, in run order:\n\n'
    for s in "${SETUPS[@]}"; do printf '  %s\n' "$(basename "$s")"; done
    printf '\n  %d setup(s), resolved from %s\n' "${#SETUPS[@]}" "$SETUP_DIR"
    printf '  records per worksheet: %s\n\n' "$RECORDS"
    exit 0
fi

if [ "${#SETUPS[@]}" -eq 0 ]; then
    printf '%sverify-package.sh: no setups under %s%s\n' "$C_SKIP" "$SETUP_DIR" "$C_OFF"
    exit 2
fi

# --- the run dir -------------------------------------------------------------
STAMP="$(date +%Y%m%d-%H%M%S)"
export VERIFY_WORK="$REPO/.obt/draft/verify-package/run-$STAMP"
mkdir -p "$VERIFY_WORK"
RESULTS="$VERIFY_WORK/results.txt"

# shellcheck source=/dev/null
source "$VERIFY_LIB"

printf 'verify-package.sh: run %s\n' "$STAMP"
printf 'verify-package.sh: work %s\n' "$VERIFY_WORK"
printf 'verify-package.sh: setups %s\n' "$SETUP_DIR"

# --- gate 1: is the installed obt the latest? --------------------------------
printf '\n%s\n' "──── gate: the obt package ────"
PKG_LOG="$VERIFY_WORK/gate-package-version.log"
R_PROFILE_USER=nonexistent Rscript "$PKG_DIR/gates/package-version.R" \
    > "$PKG_LOG" 2>&1
sed 's/^/  /' "$PKG_LOG"

pkg_value() { sed -n "s/^$1=//p" "$PKG_LOG" | head -1; }
PKG_CURRENT="$(pkg_value PKG_CURRENT)"
PKG_WHY="$(pkg_value PKG_WHY)"
PKG_REINSTALL="$(pkg_value PKG_REINSTALL)"

if [ "$PKG_CURRENT" != "yes" ]; then
    printf '\n  %sSKIP%s the installed obt is not the one to judge: %s\n' \
        "$C_SKIP" "$C_OFF" "${PKG_WHY:-unknown}"
    [ -n "$PKG_REINSTALL" ] && printf '  install it with: %s\n' "$PKG_REINSTALL"
    printf '  then re-run. Nothing was generated, so nothing was cleaned up.\n'
    exit 2
fi

# --- gate 2: can obt drive the cached channel? -------------------------------
printf '\n%s\n' "──── gate: the obt channel ────"
CHANNEL_SRC=""
for candidate in \
    "${OBT_VERIFY_CHANNEL:-}" \
    "$REPO/.obt/releases/latest/OBT-dev-latest" \
    "$REPO/releases/latest/OBT-dev-latest"
do
    if [ -n "$candidate" ] && [ -d "$candidate" ] && ls "$candidate"/*.zip >/dev/null 2>&1; then
        CHANNEL_SRC="$candidate"
        break
    fi
done

if [ -z "$CHANNEL_SRC" ]; then
    printf '  %sSKIP%s no channel zip cached on this machine.\n' "$C_SKIP" "$C_OFF"
    printf '  Build one: bash scripts/release/build-release-zip.sh dev latest\n'
    printf '  or point OBT_VERIFY_CHANNEL at a folder holding one.\n'
    exit 2
fi
printf '  cached channel: %s\n' "$CHANNEL_SRC"

CH_LOG="$VERIFY_WORK/gate-channel-age.log"
R_PROFILE_USER=nonexistent Rscript "$PKG_DIR/gates/channel-age.R" \
    "$CHANNEL_SRC" > "$CH_LOG" 2>&1
sed 's/^/  /' "$CH_LOG"

ch_value() { sed -n "s/^$1=//p" "$CH_LOG" | head -1; }
CH_DRIVABLE="$(ch_value CHANNEL_DRIVABLE)"
CH_WHY="$(ch_value CHANNEL_WHY)"
CH_VERSION="$(ch_value CHANNEL_VERSION)"
CH_NEWEST_EPOCH="$(ch_value CHANNEL_NEWEST_EPOCH)"
CH_NEWEST_DATE="$(ch_value CHANNEL_NEWEST_DATE)"

if [ "$CH_DRIVABLE" != "yes" ]; then
    printf '\n  %sSKIP%s obt cannot drive the cached channel: %s\n' \
        "$C_SKIP" "$C_OFF" "${CH_WHY:-unreadable version}"
    printf '  Build a current one: bash scripts/release/build-release-zip.sh dev latest\n'
    printf '  then re-run. Nothing was generated, so nothing was cleaned up.\n'
    exit 2
fi

# The lag is printed, never a gate. A drivable channel is worth running even
# when it trails src/; what it must not do is trail silently.
CHANNEL_NOTE=""
SRC_EPOCH="$(git -C "$REPO" log -1 --format=%ct -- src/ 2>/dev/null)"
if [ -n "$CH_NEWEST_EPOCH" ] && [ -n "$SRC_EPOCH" ]; then
    LAG_DAYS=$(( (SRC_EPOCH - CH_NEWEST_EPOCH) / 86400 ))
    if [ "$LAG_DAYS" -gt 0 ]; then
        printf '  the channel carries %s, %s day(s) behind src/\n' "$CH_NEWEST_DATE" "$LAG_DAYS"
        CHANNEL_NOTE=" (channel ${LAG_DAYS}d behind src/)"
    else
        printf '  the channel is level with src/\n'
    fi
fi

# --- the watchdog ------------------------------------------------------------
GUARD_TAG="verifypkg-$STAMP"
GUARD_SCRIPT="$REPO/.obt/draft/headless/guard-start.ps1"
GUARD_UP=0

guard_up() {
    [ "$NO_GUARD" -eq 1 ] && { printf '\nverify-package.sh: watchdog OFF (--no-guard)\n'; return 0; }
    if [ ! -f "$GUARD_SCRIPT" ]; then
        printf '\nverify-package.sh: no guard-start.ps1, running unwatched\n'
        return 0
    fi
    if powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$GUARD_SCRIPT" \
            -Tag "$GUARD_TAG" -DeadlineSeconds 7200 -FloorRamGB 1.2 2>&1 | sed 's/^/  /'; then
        GUARD_UP=1
    else
        printf '%sverify-package.sh: the watchdog would not start.%s An Excel is probably open.\n' \
            "$C_FAIL" "$C_OFF"
        printf '  Close it, or re-run with --no-guard and accept that nothing watches Excel.\n'
        exit 2
    fi
}

guard_down() {
    [ "$GUARD_UP" -eq 1 ] || return 0
    powershell.exe -NoProfile -ExecutionPolicy Bypass \
        -File "$REPO/.obt/draft/hidden-instance-probe/guard-ctl.ps1" \
        -Action stop -Tag "$GUARD_TAG" 2>&1 | sed 's/^/  /' | head -6
}
trap guard_down EXIT

printf '\n'
guard_up

# --- run them ----------------------------------------------------------------
PASSED=0; FAILED=0; SKIPPED=0
printf '\n'

for setup in "${SETUPS[@]}"; do
    stem="$(basename "$setup")"; stem="${stem%.*}"
    if [ -n "$PATTERN" ] && ! printf '%s' "$stem" | grep -q -- "$PATTERN"; then
        continue
    fi

    printf '%s\n' "────────────────────────────────────────────────────────────────"
    printf '%s  through the obt package\n' "$stem"

    VERIFY_FAILURES=0
    work="$VERIFY_WORK/$stem"
    log="$VERIFY_WORK/$stem.log"
    started=$(date +%s)

    R_PROFILE_USER=nonexistent Rscript "$PKG_DIR/run-setup.R" \
        "$work" "$setup" "$CHANNEL_SRC" "ll-$stem" "$RECORDS" \
        > "$log" 2>&1
    driver_code=$?
    sed 's/^/      /' "$log"

    value() { sed -n "s/^$1=//p" "$log" | head -1; }
    LINELIST="$(value PKG_LINELIST)"
    FAKE="$(value PKG_FAKE)"
    IMPORTED="$(value PKG_IMPORTED)"
    FAILED_STAGE="$(value PKG_STAGE_FAILED)"

    if [ "$driver_code" -ne 0 ]; then
        say "the run stopped at stage: ${FAILED_STAGE:-unknown}"
        say "its message: $(value PKG_ERROR)"
    fi

    # JUDGE THE OUTPUT, NOT THE EXIT CODE. The driver answering 0 over a
    # linelist that will not open is the thing this suite exists to catch.
    assert_equal "$driver_code" "0" "the obt recipe committed without raising"
    assert_file "$LINELIST" "obt generated a linelist"
    if [ -n "$LINELIST" ] && [ -f "$LINELIST" ]; then
        assert_file_bigger_than "$LINELIST" 1000000 "the linelist is a real workbook"
        say "delivered a ${LINELIST##*.} linelist"
    fi
    assert_file "$FAKE" "obt_fake wrote a dataset drawn from the setup"
    assert_equal "$IMPORTED" "yes" "the fake dataset imported into the generated linelist"

    # The same judge the verify suite uses, so "coherent" means one thing
    # whichever suite says it.
    if [ -n "$LINELIST" ] && [ -f "$LINELIST" ]; then
        inspect_linelist "$LINELIST" "$VERIFY_WORK/$stem-inspect.log"
        assert $? "linelist-inspect.R reports the delivered linelist clean"
        if [ -s "${VERIFY_LAST_INSPECT_LOG:-}" ]; then
            grep -iE "fail|missing|absent" "${VERIFY_LAST_INSPECT_LOG}" | head -6 | sed 's/^/      /'
        fi
    fi

    elapsed=$(( $(date +%s) - started ))
    if [ "$VERIFY_FAILURES" -eq 0 ]; then
        verdict="PASS"; PASSED=$((PASSED + 1))
        reason="generated, faked and imported$CHANNEL_NOTE"
        printf '    %sPASS%s %s\n' "$C_OK" "$C_OFF" "$reason"
    else
        verdict="FAIL"; FAILED=$((FAILED + 1))
        reason="$VERIFY_FAILURES assertion(s) did not hold${FAILED_STAGE:+, stopped at $FAILED_STAGE}"
        printf '    %sFAIL%s %s\n' "$C_FAIL" "$C_OFF" "$reason"
    fi
    printf '%-40s %-5s %4ss  %s\n' "$stem" "$verdict" "$elapsed" "$reason" >> "$RESULTS"
done

# --- the table ---------------------------------------------------------------
printf '\n%s\n' "════════════════════════════════════════════════════════════════"
printf '%s\n\n' "verify-package.sh: results"
if [ -f "$RESULTS" ]; then
    while IFS= read -r line; do
        case "$line" in
            *" PASS "*) printf '  %s%s%s\n' "$C_OK" "$line" "$C_OFF" ;;
            *" FAIL "*) printf '  %s%s%s\n' "$C_FAIL" "$line" "$C_OFF" ;;
            *" SKIP "*) printf '  %s%s%s\n' "$C_SKIP" "$line" "$C_OFF" ;;
            *)          printf '  %s\n' "$line" ;;
        esac
    done < "$RESULTS"
fi
printf '\n  %d passed, %d failed, %d skipped\n' "$PASSED" "$FAILED" "$SKIPPED"
printf '  obt %s, channel %s\n' "$(pkg_value PKG_INSTALLED_VERSION)" "${CH_VERSION:-?}"
printf '  logs and generated files: %s\n' "$VERIFY_WORK"

# --- clean up after itself ---------------------------------------------------
# Every per-setup working tree goes, logs stay. A tree holds a copied channel
# zip, a setup, a linelist and a fake dataset, so leaving six of them behind
# costs well over a hundred megabytes on a laptop that has none to spare.
if [ "$KEEP" -eq 1 ]; then
    printf '  --keep: every generated file was kept\n'
elif [ "$FAILED" -ne 0 ]; then
    printf '  %sfailed run: every generated file was kept%s\n' "$C_FAIL" "$C_OFF"
else
    freed=0
    if command -v du >/dev/null 2>&1; then
        freed="$(du -sm "$VERIFY_WORK" 2>/dev/null | cut -f1)"
    fi
    for setup in "${SETUPS[@]}"; do
        stem="$(basename "$setup")"; stem="${stem%.*}"
        [ -d "$VERIFY_WORK/$stem" ] && rm -rf "$VERIFY_WORK/$stem"
    done
    printf '  generated trees deleted (%sMB), logs kept (--keep to keep both)\n' "${freed:-?}"
fi

[ "$FAILED" -eq 0 ] || exit 1
exit 0

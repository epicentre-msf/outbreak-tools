#!/usr/bin/env bash
#
# verify.sh [--list] [--case <pattern>] [--keep] [--no-guard] [--help]
#
# The behavioural suite: generate real files with real Excel and judge whether
# what came out is coherent. Run it before a release.
#
#   bash scripts/verify/verify.sh                run every case
#   bash scripts/verify/verify.sh --list         name them, run nothing
#   bash scripts/verify/verify.sh --case 030     run the ones matching 030
#   bash scripts/verify/verify.sh --keep         keep the generated files
#
# WHAT THIS IS, AGAINST WHAT ELSE THERE IS
# -----------------------------------------------------------------------------
#   scripts/tests/run-tests.R      the UNIT suites, inside Excel, over src/
#   scripts/release/preflight.sh   the STATIC release gate: envelopes, ribbons,
#                                  hand-imported drift, closure, translations
#   scripts/verify/verify.sh       THIS: the product's BEHAVIOUR, end to end
#
# The preflight asks whether the binaries ship correctly. This asks whether the
# thing they build is any good. A release wants both, and neither answers the
# other's question.
#
# Exit 0  every case passed or skipped
# Exit 1  at least one case failed
# Exit 2  the suite could not run at all
#
# EXCEL, AND WHY THE SUITE MAY REFUSE TO START
# -----------------------------------------------------------------------------
# Every case launches its own Excel through the headless build. A watchdog
# stops a build that runs away on memory, and it will not start while another
# Excel is open, because it cannot tell a runaway build from somebody's own
# workbook. If you have Excel open, close it or pass --no-guard and accept that
# nothing is watching. On a release run, close Excel.
set -uo pipefail

cd "$(git rev-parse --show-toplevel 2>/dev/null || echo .)"
REPO="$(pwd)"
VERIFY_DIR="$REPO/scripts/verify"
CASE_DIR="$VERIFY_DIR/cases"

LIST_ONLY=0
KEEP=0
NO_GUARD=0
PATTERN=""

while [ $# -gt 0 ]; do
    case "$1" in
        --list)     LIST_ONLY=1 ;;
        --keep)     KEEP=1 ;;
        --no-guard) NO_GUARD=1 ;;
        --case)     shift; PATTERN="${1:-}" ;;
        --case=*)   PATTERN="${1#--case=}" ;;
        -h|--help)  sed -n '2,33p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
        *)          printf 'verify.sh: unknown option %s\n' "$1" >&2; exit 2 ;;
    esac
    shift
done

if [ -t 1 ] && [ -z "${NO_COLOR:-}" ]; then
    C_FAIL=$'\033[1;31m'; C_OK=$'\033[32m'; C_SKIP=$'\033[33m'; C_OFF=$'\033[0m'
else
    C_FAIL=""; C_OK=""; C_SKIP=""; C_OFF=""
fi

# --- the cases ---------------------------------------------------------------
mapfile -t CASES < <(find "$CASE_DIR" -maxdepth 1 -name '*.sh' -type f 2>/dev/null | sort)
if [ "${#CASES[@]}" -eq 0 ]; then
    printf 'verify.sh: no case under %s\n' "$CASE_DIR" >&2
    exit 2
fi

# The declarative header of a case: "# title:" and "# asserts:" lines.
case_field() {
    sed -n "s/^# $2: *//p" "$1" 2>/dev/null | head -1
}

if [ "$LIST_ONLY" -eq 1 ]; then
    printf '%s\n' "The verify cases, in run order:"
    printf '\n'
    for c in "${CASES[@]}"; do
        id="$(basename "$c" .sh)"
        printf '  %-44s %s\n' "$id" "$(case_field "$c" title)"
        a="$(case_field "$c" asserts)"
        [ -n "$a" ] && printf '  %-44s   asserts: %s\n' "" "$a"
    done
    printf '\n  %d case(s). Fixtures resolve under %s\n' "${#CASES[@]}" \
        "${OBT_VERIFY_FIXTURES:-.obt/draft/headless}"
    exit 0
fi

# --- the run dir -------------------------------------------------------------
STAMP="$(date +%Y%m%d-%H%M%S)"
export VERIFY_WORK="$REPO/.obt/draft/verify/run-$STAMP"
mkdir -p "$VERIFY_WORK"
RESULTS="$VERIFY_WORK/results.txt"

printf 'verify.sh: run %s\n' "$STAMP"
printf 'verify.sh: work %s\n' "$VERIFY_WORK"
printf 'verify.sh: fixtures %s\n' "${OBT_VERIFY_FIXTURES:-$REPO/.obt/draft/headless}"

# --- the watchdog ------------------------------------------------------------
# One watchdog covers the whole run. guard-start.ps1 is the headless folder's
# wrapper: it starts the same watchdog guard-ctl.ps1 does and skips only
# guard-ctl's 10GB free-commit pre-flight, which never passes on the owner's
# laptop. The cap, the RAM floor and the deadline all stay.
GUARD_TAG="verify-$STAMP"
GUARD_SCRIPT="$REPO/.obt/draft/headless/guard-start.ps1"
GUARD_UP=0

guard_up() {
    [ "$NO_GUARD" -eq 1 ] && { printf 'verify.sh: watchdog OFF (--no-guard)\n'; return 0; }
    if [ ! -f "$GUARD_SCRIPT" ]; then
        printf 'verify.sh: no guard-start.ps1, running unwatched\n'
        return 0
    fi
    if powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$GUARD_SCRIPT" \
            -Tag "$GUARD_TAG" -DeadlineSeconds 7200 -FloorRamGB 1.2 2>&1 | sed 's/^/  /'; then
        GUARD_UP=1
    else
        printf '%sverify.sh: the watchdog would not start.%s An Excel is probably open.\n' \
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

guard_up

# --- run them ----------------------------------------------------------------
PASSED=0; FAILED=0; SKIPPED=0
printf '\n'

for c in "${CASES[@]}"; do
    id="$(basename "$c" .sh)"
    if [ -n "$PATTERN" ] && ! printf '%s' "$id" | grep -q -- "$PATTERN"; then
        continue
    fi

    title="$(case_field "$c" title)"
    log="$VERIFY_WORK/$id.log"

    printf '%s\n' "────────────────────────────────────────────────────────────────"
    printf '%s\n' "$id  $title"

    started=$(date +%s)
    VERIFY_WORK="$VERIFY_WORK" bash "$c" 2>&1 | tee "$log" | sed 's/^/  /'
    code="${PIPESTATUS[0]}"
    elapsed=$(( $(date +%s) - started ))

    case "$code" in
        0) verdict="PASS"; PASSED=$((PASSED + 1)) ;;
        2) verdict="SKIP"; SKIPPED=$((SKIPPED + 1)) ;;
        *) verdict="FAIL"; FAILED=$((FAILED + 1)) ;;
    esac

    # The reason line the case ended on, for the table.
    reason="$(grep -Eo '(PASS|SKIP|FAIL)[[:space:]]+.*' "$log" | tail -1 | sed 's/^[A-Z]* *//')"
    printf '%-44s %-5s %4ss  %s\n' "$id" "$verdict" "$elapsed" "${reason:-}" >> "$RESULTS"
done

# --- the table ---------------------------------------------------------------
printf '\n%s\n' "════════════════════════════════════════════════════════════════"
printf '%s\n\n' "verify.sh: results"
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
printf '  logs and generated files: %s\n' "$VERIFY_WORK"

if [ "$KEEP" -eq 0 ] && [ "$FAILED" -eq 0 ]; then
    # The workbooks go and the logs stay: a green run's 40MB of linelists is
    # rebuilt by the next one, and its logs are the record that it was green.
    find "$VERIFY_WORK" -type f \( -name '*.xlsm' -o -name '*.xlsb' \) -delete 2>/dev/null
    printf '  workbooks cleared, logs kept (--keep to keep both)\n'
elif [ "$FAILED" -ne 0 ]; then
    printf '  %sfailed run: every generated file was kept%s\n' "$C_FAIL" "$C_OFF"
fi

[ "$FAILED" -eq 0 ] || exit 1
exit 0

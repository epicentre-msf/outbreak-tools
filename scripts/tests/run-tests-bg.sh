#!/usr/bin/env bash
#
# run-tests-bg.sh — launch the headless test loop and ping once it is done.
#
# The loop takes several minutes and its stdout is thousands of inventory
# lines, so a caller that waits on it in the foreground either times out or
# drowns. This wrapper detaches the run, keeps the whole log on disk, and
# prints only the few lines worth reading.
#
# Usage:
#   scripts/tests/run-tests-bg.sh [--build] [other run-tests.R flags]
#   scripts/tests/run-tests-bg.sh --wait          # block until the run ends
#
# Pass --build unless you know the workbook tables are current: without it the
# run imports nothing and reports modulesRun=0.
#
# Where things land:
#   .test-runner/bg/run.log      the full runner output
#   .test-runner/bg/summary.txt  the summary line and any failure rows
#   .test-runner/bg/status       the exit code, written last
#
# On finishing it posts a macOS notification and prints the summary.

set -uo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
OUT_DIR="$REPO_ROOT/.test-runner/bg"
LOG_FILE="$OUT_DIR/run.log"
SUMMARY_FILE="$OUT_DIR/summary.txt"
STATUS_FILE="$OUT_DIR/status"
RESULTS_CSV="$REPO_ROOT/.test-runner/tests/staging/test-results.csv"

WAIT_FOR_RUN=0
RUNNER_ARGS=()
for arg in "$@"; do
    if [[ "$arg" == "--wait" ]]; then
        WAIT_FOR_RUN=1
    else
        RUNNER_ARGS+=("$arg")
    fi
done

mkdir -p "$OUT_DIR"
rm -f "$STATUS_FILE" "$SUMMARY_FILE"

# --- helpers ---------------------------------------------------------------

notify() {
    # A notification is best-effort: a machine with notifications switched off
    # must not fail the run.
    osascript -e "display notification \"$2\" with title \"$1\"" >/dev/null 2>&1 || true
}

quit_excel() {
    # A live Excel wedges the AppleScript trigger, and the runner reports that
    # as the opaque -50 or -609. Quitting first is the cheapest fix there is.
    pgrep -x "Microsoft Excel" >/dev/null 2>&1 || return 0

    osascript -e 'with timeout of 30 seconds
tell application "Microsoft Excel" to quit saving no
end timeout' >/dev/null 2>&1

    local waited=0
    while pgrep -x "Microsoft Excel" >/dev/null 2>&1 && [[ $waited -lt 20 ]]; do
        sleep 1
        waited=$((waited + 1))
    done

    if pgrep -x "Microsoft Excel" >/dev/null 2>&1; then
        echo "run-tests-bg: Excel is still up and is not answering AppleEvents." >&2
        echo "run-tests-bg: a dialog is probably open. Clear it, then run this again." >&2
        return 1
    fi
}

# --- the run itself --------------------------------------------------------

run_suite() {
    quit_excel || {
        echo "127" >"$STATUS_FILE"
        printf 'Excel would not quit; nothing was run.\n' >"$SUMMARY_FILE"
        notify "OutbreakTools tests" "Blocked: Excel would not quit"
        return 127
    }

    Rscript "$REPO_ROOT/scripts/tests/run-tests.R" "${RUNNER_ARGS[@]}" >"$LOG_FILE" 2>&1
    local code=$?

    # The inventory is one line per component and drowns everything else, so
    # the summary keeps the runner's own report lines and drops those.
    {
        grep -vE '^ *comp:' "$LOG_FILE" \
            | grep -E 'summary —|success / |latest results|all tests passed|\[FAIL\]|execution error|error:' \
            || echo "No summary line was written. Read $LOG_FILE."

        if [[ -f "$RESULTS_CSV" ]]; then
            # grep -c exits 1 on a zero count, so the fallback goes on the
            # assignment. Written as `|| echo 0` inside the substitution it
            # appends a second line and every clean run reads as a failure.
            local failures
            failures=$(grep -c '"failure"' "$RESULTS_CSV" 2>/dev/null) || failures=0
            echo "failure rows in test-results.csv: $failures"
            if [[ "$failures" != "0" ]]; then
                echo "--- failures ---"
                grep '"failure"' "$RESULTS_CSV" | head -40
            fi
        fi
    } >"$SUMMARY_FILE" 2>&1

    echo "$code" >"$STATUS_FILE"

    local headline
    headline=$(grep -E 'summary —' "$SUMMARY_FILE" | head -1)
    [[ -n "$headline" ]] || headline="exit code $code"

    if [[ $code -eq 0 ]]; then
        notify "OutbreakTools tests passed" "$headline"
    else
        notify "OutbreakTools tests FAILED" "$headline"
    fi

    return $code
}

if [[ $WAIT_FOR_RUN -eq 1 ]]; then
    run_suite
    code=$?
    cat "$SUMMARY_FILE"
    exit $code
fi

# Detached: the caller gets the pid straight away and reads the summary when
# the status file appears.
run_suite >/dev/null 2>&1 &
RUN_PID=$!

cat <<EOF
run-tests-bg: launched (pid $RUN_PID)
  log      $LOG_FILE
  summary  $SUMMARY_FILE
  status   $STATUS_FILE   (written last; its content is the exit code)

Wait on it with:
  while [ ! -f "$STATUS_FILE" ]; do sleep 15; done; cat "$SUMMARY_FILE"
EOF

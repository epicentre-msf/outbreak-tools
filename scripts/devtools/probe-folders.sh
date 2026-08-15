#!/usr/bin/env bash
# probe-folders.sh [folder ...]
#
# Walk the test registry one suite folder at a time.
#
# For each folder: narrow the registry to it with narrow-registry.R, run
# `run-tests.R --build`, log the run, then restore the registry with
# `git checkout --` so the file is byte-identical for any other session reading
# it. The loop CONTINUES past a red folder, so one pass gives the whole picture
# rather than stopping at the first fault.
#
# Why folder by folder rather than one full run: a full run is around half an
# hour and a wedged Excel costs all of it, while a folder probe is short enough
# to repeat. A red folder also names itself, so there is nothing to bisect.
#
# With no arguments it walks every folder the registry declares except
# `helpers` -- that block carries no suite of its own, so probing it alone
# would run 0 tests, which the runner reports as a failure.
#
# Excel is quit between runs: `run VB macro` fails against an already-open
# Excel, and only one test run may be in flight at a time.
#
# Logs land in .test-runner/probe-logs/<folder>.log with a summary.txt beside
# them. That folder is gitignored with the rest of .test-runner.
set -u

REPO=$(cd -- "$(dirname -- "$0")/../.." && pwd)
cd "$REPO" || exit 1

LOGS="${OBT_TEST_HOME:-$REPO/.test-runner}/probe-logs"
mkdir -p "$LOGS"

if [ "$#" -gt 0 ]; then
  FOLDERS=("$@")
else
  # Every `- folder:` in the registry, minus helpers. Read in a loop rather
  # than with mapfile: macOS ships bash 3.2, which does not have it.
  FOLDERS=()
  while IFS= read -r name; do
    FOLDERS+=("$name")
  done < <(sed -n 's/^  - folder: //p' src/tests/test-registry.yml | grep -v '^helpers$')
fi

SUMMARY="$LOGS/summary.txt"
: > "$SUMMARY"
red=0

for f in "${FOLDERS[@]}"; do
  echo "=== $f — starting $(date '+%H:%M:%S') ===" | tee -a "$SUMMARY"
  git checkout -- src/tests/test-registry.yml
  Rscript scripts/devtools/narrow-registry.R "$f" > "$LOGS/$f.narrow.txt" 2>&1 || {
    echo "$f: SKIPPED — narrowing failed, see $LOGS/$f.narrow.txt" | tee -a "$SUMMARY"
    continue
  }
  pkill -x "Microsoft Excel" 2>/dev/null
  sleep 3

  Rscript scripts/tests/run-tests.R --build > "$LOGS/$f.log" 2>&1
  rc=$?
  git checkout -- src/tests/test-registry.yml

  counts=$(grep -m1 "success / " "$LOGS/$f.log")
  summary_line=$(grep -m1 "summary —" "$LOGS/$f.log")
  if [ "$rc" -eq 0 ]; then
    echo "$f: GREEN | $counts | $summary_line" | tee -a "$SUMMARY"
  else
    red=$((red + 1))
    echo "$f: RED rc=$rc | $counts | $summary_line" | tee -a "$SUMMARY"
    grep -A60 "^Failing rows:" "$LOGS/$f.log" | tee -a "$SUMMARY"
  fi
  echo "" | tee -a "$SUMMARY"
done

pkill -x "Microsoft Excel" 2>/dev/null
echo "=== done $(date '+%H:%M:%S'), $red red folder(s). Logs: $LOGS ===" | tee -a "$SUMMARY"
exit "$red"

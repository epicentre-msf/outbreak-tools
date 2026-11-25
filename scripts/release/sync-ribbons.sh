#!/usr/bin/env bash
#
# sync-ribbons.sh [--check] [--no-bin]
#
# Release guard against ribbon drift. Every workbook push-assets.sh bundles
# carries a ribbon that also lives, unpacked, under ribbons/. The two must agree
# byte for byte, or a build ships a ribbon nobody can see in the sources — and
# the next ribbon-pack.sh silently reverts whichever side was ahead.
#
# So, before any binary leaves this machine:
#
#   1. compare each workbook's customUI part and its images against the folder
#      that holds them. Any difference is printed and the run STOPS.
#   2. once every pair agrees, rebuild those folders out of the workbooks, so
#      the tracked sources are a straight readout of what ships. Step 1 makes
#      this a no-op on content; it is what drops a stray file or a renamed icon.
#
# --check runs step 1 only.
#
# --no-bin leaves every workbook under src/bin out of both steps. It is for a
# push that sends the mocks and the ribbon templates alone
# (push-assets.sh --mock-ribbons): the dev builds stay on this disk, so one that
# trails its mock stops nothing.
#
# The verdict is byte-exact — that is the whole point — but a byte diff reads
# badly, because a workbook's customUI can mix CRLF and LF line endings and GNU
# diff then calls every line changed. So the check decides on bytes and reports
# with --strip-trailing-cr, saying so when the endings are the only difference.
#
# .source is not compared and not rewritten: it records the workbook
# ribbon-pack.sh writes back into by default, which is not always the workbook a
# folder is checked against (ribbons/setup tracks the dev line but packs into the
# main workbook).
#
# The pairing below is explicit rather than read out of .source, because two
# workbooks share one folder: .mock/setup_mock.xlsm and src/bin/setup/setup_dev.xlsm
# are two copies of the one ribbon in ribbons/setup_mock, and .source can name only
# one of them.
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"

# workbook|ribbons folder. One folder per ribbon, named after the mock that owns
# it, and both workbooks carrying that ribbon are checked against it: the mock is
# the authority and the dev twin is a copy of it (scripts/devtools/update-files.R),
# so a dev build that trails its mock cannot slip into the bundle.
#
# The promoted main builds (designer, setup, msetup) and the main ribbon template
# have no folder and are not checked. They are promoted from the dev line, so they
# trail it by design and would fail this check every time.
#
# The workbooks are the .xlsm ones. An .xlsb twin left on disk is no longer a
# ribbon source: it keeps whatever ribbon it had when the .xlsm was saved from it.
PAIRS=(
  # designer
  ".mock/designer_mock.xlsm|ribbons/designer_mock"
  "src/bin/designer/designer_dev.xlsm|ribbons/designer_mock"
  # setup
  ".mock/setup_mock.xlsm|ribbons/setup_mock"
  "src/bin/setup/setup_dev.xlsm|ribbons/setup_mock"
  # unit tests
  ".mock/unit_tests_mock.xlsm|ribbons/unit_tests_mock"
  "src/bin/test-files/unit_tests_dev.xlsm|ribbons/unit_tests_mock"
  # master setup
  ".mock/mastersetup_mock.xlsm|ribbons/mastersetup_mock"
  "src/bin/msetup/msetup_dev.xlsm|ribbons/mastersetup_mock"
  # ribbon template
  "ribbons/_ribbontemplate_dev.xlsm|ribbons/_ribbontemplate_dev"
)
# .mock/cleanDesignerMockFile is deliberately absent: it is a stripped designer
# kept for rebuilds, not one of the shipped ribbon sources.

CHECK_ONLY=0
NO_BIN=0
for arg in "$@"; do
  case "$arg" in
    --check)   CHECK_ONLY=1 ;;
    --no-bin)  NO_BIN=1 ;;
    -h|--help) sed -n '2,38p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *)         echo "unknown option: $arg" >&2; exit 2 ;;
  esac
done

SCOPE_NOTE=""
if [ $NO_BIN -eq 1 ]; then
  KEPT=()
  for pair in "${PAIRS[@]}"; do
    case "$pair" in
      src/bin/*) ;;
      *)         KEPT+=("$pair") ;;
    esac
  done
  PAIRS=("${KEPT[@]}")
  SCOPE_NOTE=" (src/bin left out)"
fi

for t in unzip diff; do
  command -v "$t" >/dev/null 2>&1 || { echo "ERROR: $t not found." >&2; exit 1; }
done
EXTRACT="scripts/devtools/ribbon-extract.sh"
[ -f "$EXTRACT" ] || { echo "ERROR: $EXTRACT is missing." >&2; exit 1; }

TMPD="$(mktemp -d)"; trap 'rm -rf "$TMPD"' EXIT

echo "==> ribbons vs their unpacked sources$SCOPE_NOTE"
FAILED=0
for pair in "${PAIRS[@]}"; do
  WB="${pair%%|*}"; DIR="${pair#*|}"
  STEM=$(basename "$WB"); STEM="${STEM%.*}"

  if [ ! -f "$WB" ]; then
    if [ -f "${WB%.*}.xlsb" ]; then
      echo "  FAIL  $WB is missing; only ${WB%.*}.xlsb is there (RELEASING.md §2: making the .xlsm binaries)" >&2
    else
      echo "  FAIL  $WB is missing (pull-assets.sh first?)" >&2
    fi
    FAILED=1; continue
  fi
  if [ ! -d "$DIR" ]; then
    echo "  FAIL  $DIR is missing — it is the ribbon source for $WB" >&2; FAILED=1; continue
  fi

  if ! bash "$EXTRACT" "$WB" --out-dir "$TMPD" --overwrite >"$TMPD/$STEM.log" 2>&1; then
    echo "  FAIL  could not read a ribbon out of $WB:" >&2
    sed 's/^/          /' "$TMPD/$STEM.log" >&2
    FAILED=1; continue
  fi

  # byte-exact verdict, over ribbon.xml and every icon
  if diff -rq --exclude=.source "$TMPD/$STEM" "$DIR" >"$TMPD/$STEM.brief" 2>&1; then
    echo "  ok    $WB  ==  $DIR"
    continue
  fi

  echo "  FAIL  $WB  !=  $DIR" >&2
  sed 's/^/          /' "$TMPD/$STEM.brief" >&2
  if [ -f "$TMPD/$STEM/ribbon.xml" ] && [ -f "$DIR/ribbon.xml" ]; then
    diff -u --strip-trailing-cr --label "in $WB" --label "$DIR/ribbon.xml" \
         "$TMPD/$STEM/ribbon.xml" "$DIR/ribbon.xml" >"$TMPD/$STEM.diff" 2>&1 || true
    if [ -s "$TMPD/$STEM.diff" ]; then
      head -60 "$TMPD/$STEM.diff" | sed 's/^/          /' >&2
      [ "$(wc -l <"$TMPD/$STEM.diff")" -gt 60 ] && echo "          ... diff truncated" >&2
    else
      echo "          (the XML matches line for line — the line endings differ)" >&2
    fi
  fi
  FAILED=1
done

if [ $FAILED -ne 0 ]; then
  cat >&2 <<'MSG'

ERROR: the shipped ribbons and the sources under ribbons/ have drifted apart.
       Nothing was pushed. Decide which side is right, then make them agree:

         # the folder is ahead (a ribbon edited as text):
         scripts/devtools/ribbon-pack.sh <folder> --out <workbook> --overwrite
         scripts/devtools/ribbon-doctor.sh <workbook>

         # the workbook is ahead (a ribbon edited in Excel / OfficeRibbonX):
         scripts/devtools/ribbon-extract.sh <workbook> --out-dir ribbons --overwrite

       A dev workbook trailing its mock is not a ribbon edit at all — run the
       'update ... dev' tasks (scripts/devtools/update-files.R) to recopy it.
MSG
  exit 1
fi

[ $CHECK_ONLY -eq 1 ] && exit 0

echo "==> rebuilding the mock and dev ribbon sources from the workbooks$SCOPE_NOTE"
for pair in "${PAIRS[@]}"; do
  WB="${pair%%|*}"; DIR="${pair#*|}"
  STEM=$(basename "$WB"); STEM="${STEM%.*}"

  # .source is carried across as a file, not as text: reading and rewriting it
  # would drop its trailing newline and show up as a diff in every rebuild.
  KEEP=""
  if [ -f "$DIR/.source" ]; then KEEP="$TMPD/keep.$(basename "$DIR")"; cp "$DIR/.source" "$KEEP"; fi
  rm -rf "$DIR"
  cp -R "$TMPD/$STEM" "$DIR"
  rm -f "$DIR/.source"
  [ -n "$KEEP" ] && mv "$KEEP" "$DIR/.source"
  echo "  $WB -> $DIR"
done

echo "Ribbons are in sync."

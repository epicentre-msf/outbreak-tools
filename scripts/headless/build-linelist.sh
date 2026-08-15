#!/usr/bin/env bash
# build-linelist.sh
# =============================================================================
# Generate one linelist, headless. A wrapper over build-linelist.R that holds
# the settings for a run in one place instead of a long command line.
#
#   ./scripts/headless/build-linelist.sh                 build with the settings below
#   ./scripts/headless/build-linelist.sh --help          the full parameter list
#   ./scripts/headless/build-linelist.sh --setup=x.xlsb  override any setting
#
# Anything passed on the command line is handed straight to build-linelist.R
# and wins over the settings below, so a saved default never has to be edited
# for a one-off run.
#
# Exit 0  a linelist was written and copied to the output folder
# Exit 1  anything else, with the build's own account printed
# =============================================================================

set -euo pipefail

# --- the settings, edit these ------------------------------------------------
# Paths may be relative to the repo root or absolute. Leave a setting empty to
# use the build's own default, which is described against each one below.

# THE SETUP WORKBOOK. Required, and the one thing with no default. This is the
# filled setup that says which variables, sheets and choices the linelist gets.
SETUP=".obt/draft/demo_setup_filled.xlsb"

# WHERE THE FILES GO, and what they are called. The build writes
# <NAME>.xlsb, <NAME>-designer.xlsb, <NAME>-generation.txt, <NAME>-import.log
# and <NAME>-build-report.txt into this folder.
OUT=".obt/draft"
NAME="headless_linelist"

# THE RIBBON TEMPLATE. Given, the linelist gets a ribbon tab. Left empty, it
# gets action buttons on the sheets and an Admin sheet instead. The two builds
# are different products, so this is the setting most worth being sure about.
TEMPLATE="ribbons/_ribbontemplate_dev.xlsb"

# THE GEOBASE workbook, for linelists with geography. Empty means no geobase.
GEO=""

# THE DESIGNER workbook the linelist is built from. Empty uses
# .mock/designer_mock.xlsb, which is the one the project develops against.
DESIGNER=""

# THE FORMS. Empty is right for almost every run.
#
# The default build takes its forms from two FIXED folders. Neither has a flag,
# and neither is what FORMS below points at:
#
#   .mock/forms/designer        the .frm shells and their binary .frx twins
#   src/modules/linelistform    the FormLogic*.bas code merged into them
#
# The merge runs every time, so the delivered linelist has its controls wired to
# today's handlers. The result lands in <OBT_HOME>/forms.
#
# FORMS is for the other case only: a folder of ALREADY MERGED .frm files, used
# with MERGE_FORMS=no. Skipping the merge without one has nothing to build from,
# since a successful run clears its staging. And a build over unmerged forms
# ships old handlers with nothing in the file to say so.
FORMS=""
MERGE_FORMS="yes"

# LANGUAGE AND PASSWORD.
#   SETUP_LANG   which label column of the setup to read. Empty reads the
#                language the setup itself declares.
#   LL_LANG      the language of the delivered linelist, e.g. ENG. Empty keeps
#                the designer's own entry.
#   LL_PASSWORD  password applied to the linelist. Empty leaves it unprotected.
SETUP_LANG=""
LL_LANG=""
LL_PASSWORD=""

# THE STAGING ROOT. Empty is right unless something is wrong. The build stages
# itself inside Excel's own sandbox by default, which is why it runs without a
# single permission dialog:
#
#   ~/Library/Containers/com.microsoft.Excel/Data/Documents/OBTHome
#
# That needs Full Disk Access for whatever runs Rscript:
#   System Settings -> Privacy & Security -> Full Disk Access
#   add the app, then quit and reopen it.
#
# Without it the build falls back to <repo>/headless-runner, and somebody has
# to pick that folder in Excel BEFORE EVERY BUILD -- a pick covers one run and
# no more. The script prints the steps when it comes to that.
OBT_HOME=""

# WHERE THE DRIVER WORKBOOK LIVES, the untracked working area holding
# unit_tests_dev.xlsb. Empty uses <repo>/.test-runner.
BUILD_HOME=""

# --- end of settings ---------------------------------------------------------

usage() {
  cat <<'TXT'
build-linelist.sh -- generate one linelist, headless

  ./scripts/headless/build-linelist.sh [--key=value ...]

Edit the settings at the top of this file, or override any of them here.
Every flag below is passed through to build-linelist.R.

  --setup=<file>       REQUIRED. The filled setup workbook to generate from.
  --out=<dir>          Where the finished files are copied. Default: .test-runner/build
  --name=<stem>        Names every delivered file. Default: linelist

  --temppath=<file>    Ribbon template. Given -> ribbon build.
                       Omitted -> buttons build, with an Admin sheet.
  --geopath=<file>     Geobase workbook. Omitted -> no geography.
  --designer=<file>    Designer workbook. Default: .mock/designer_mock.xlsb

  --forms=<dir>        A folder of ALREADY MERGED .frm files. Only meaningful
                       with --no-merge. It is NOT where the default build reads
                       its forms from -- that is two fixed folders with no flag:
                         .mock/forms/designer      the .frm and .frx files
                         src/modules/linelistform  the code merged into them
  --no-merge           Do NOT merge src/modules/linelistform into the forms.
                       Needs --forms: a successful run clears its staging, so
                       there is no earlier copy to fall back on. Without the
                       merge the linelist ships controls wired to old handlers.

  --setuplang=<col>    Label column of the setup. Omitted -> what the setup says.
  --lllang=<code>      Linelist language, e.g. ENG. Omitted -> the designer's entry.
  --llpassword=<pass>  Password for the linelist. Omitted -> unprotected.

  --obt-home=<dir>     Staging root. Omitted -> Excel's own container, which is
                       what makes the run silent. Also read from the OBT_HOME
                       environment variable, so an .Renviron line can pin it.
  --home=<dir>         Where unit_tests_dev.xlsb lives. Default: .test-runner
  --granted            Fallback only. Records that the folder was picked in
                       Excel. Never needed inside the container.

What lands in the output folder:

  <name>.xlsb                the linelist
  <name>-designer.xlsb       the designer it was built from
  <name>-generation.txt      the build log meant for a person
  <name>-import.log          every component that went into the driver
  <name>-build-report.txt    what the trigger recorded, narrative included

Exit 0 on a delivered linelist, 1 on anything else.
TXT
}

for arg in "$@"; do
  case "$arg" in
    -h|--help) usage; exit 0 ;;
  esac
done

# Run from anywhere: every relative path in the settings is read against the
# repo root, which is also where build-linelist.R expects to resolve them.
repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

if [[ -z "$SETUP" ]] && [[ "$*" != *"--setup="* ]]; then
  echo "build-linelist.sh: no setup workbook." >&2
  echo "  Set SETUP at the top of this file, or pass --setup=<file>." >&2
  echo "  Run with --help for the full list." >&2
  exit 1
fi

# Built as an array so a path with a space in it survives. Empty settings are
# dropped rather than passed as --key=, because an empty value MEANS something
# to the build -- no ribbon template is the buttons build, no setuplang is
# "read it off the setup" -- and a default must not be mistaken for a choice.
args=()
add() { [[ -n "$2" ]] && args+=("--$1=$2"); return 0; }

add setup      "$SETUP"
add out        "$OUT"
add name       "$NAME"
add temppath   "$TEMPLATE"
add geopath    "$GEO"
add designer   "$DESIGNER"
add forms      "$FORMS"
add setuplang  "$SETUP_LANG"
add lllang     "$LL_LANG"
add llpassword "$LL_PASSWORD"
add obt-home   "$OBT_HOME"
add home       "$BUILD_HOME"

[[ "$MERGE_FORMS" == "no" ]] && args+=("--no-merge")

# The caller's flags go last so they override anything set above. Duplicate
# keys are fine: build-linelist.R reads the FIRST match, so this order would
# lose them -- hence the filter, which drops any setting the caller respecified.
if [[ $# -gt 0 ]]; then
  keep=()
  for a in "${args[@]}"; do
    key="${a%%=*}"
    dup="no"
    for given in "$@"; do
      [[ "${given%%=*}" == "$key" ]] && dup="yes" && break
    done
    [[ "$dup" == "no" ]] && keep+=("$a")
  done
  args=("${keep[@]}" "$@")
fi

echo "build-linelist.sh: Rscript scripts/headless/build-linelist.R ${args[*]}"
echo
exec Rscript scripts/headless/build-linelist.R "${args[@]}"

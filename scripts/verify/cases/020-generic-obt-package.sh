#!/usr/bin/env bash
# title: The generic setup builds through the obt package too
# asserts: obt delivers a linelist from the same setup, and linelist-inspect is clean on it
#
# WHY THIS CASE
# -----------------------------------------------------------------------------
# A user does not run build-linelist.R. They run the obt package, which drives
# clickGenerate on a downloaded designer, so it exercises a DIFFERENT path
# through the same product: Linelist.SaveLL is shared, GenerationHost's hidden
# instance is obt's and not the headless build's, and the designer is the
# released binary rather than the mock with src/ imported over it.
#
# Case 010 passing and this one failing means the release binaries disagree
# with src/, which is exactly the thing nothing in git can tell you once a
# workbook is in the store.
#
# WHY IT MAY SKIP
# -----------------------------------------------------------------------------
# obt_designer_add downloads a channel zip. This case never downloads: it
# copies a zip somebody already fetched into the working folder, where
# obt_designer_add finds and reuses it. With no cached channel anywhere, the
# case says so and skips, because a suite that reaches the network is a suite
# that goes red when GitHub does.
#
# To prepare a machine for it, fetch the channel once:
#   Rscript -e 'obt::obt_commit(obt::obt_designer_add(obt::obt(folder="<dir>"), type="dev"))'
# and it is cached under <dir>/obt/dev/ from then on. Point the case at a cache
# kept outside this repository with OBT_VERIFY_CHANNEL=<that folder>.

set -uo pipefail
source "$(dirname "${BASH_SOURCE[0]}")/../lib/verify-lib.sh"

need_command Rscript
need_file "$GENERIC_SETUP" "the generic setup"

# obt itself.
if ! R_PROFILE_USER=nonexistent Rscript -e 'quit(status = if (requireNamespace("obt", quietly = TRUE)) 0 else 1)' >/dev/null 2>&1; then
    skip_case "the obt package is not installed"
fi
say "obt is installed"

# A CACHED CHANNEL. Any folder holding a channel zip will do; the first one
# found wins. The repository's own release folders come before the obt repo's
# draft work, because they are the ones a release machine refreshes.
# OBT_VERIFY_CHANNEL points at a folder holding a channel zip somewhere else,
# for a machine that keeps one outside this repository.
CHANNEL_SRC=""
for candidate in \
    "${OBT_VERIFY_CHANNEL:-}" \
    "$REPO/.obt/releases/latest/OBT-dev-latest" \
    "$REPO/releases/latest/OBT-dev-latest"
do
    if [ -d "$candidate" ] && ls "$candidate"/*.zip >/dev/null 2>&1; then
        CHANNEL_SRC="$candidate"
        break
    fi
done

if [ -z "$CHANNEL_SRC" ]; then
    skip_case "no channel zip cached on this machine; see the header for how to fetch one"
fi
say "cached channel: $CHANNEL_SRC"
say "holding: $(ls "$CHANNEL_SRC" | tr '\n' ' ')"

WORK="$VERIFY_WORK/020-obt"
NAME="verify-generic"

step "driving the obt package"
R_PROFILE_USER=nonexistent Rscript \
    "$(dirname "${BASH_SOURCE[0]}")/020-generic-obt-package.R" \
    "$WORK" "$GENERIC_SETUP" "$CHANNEL_SRC" "$NAME" \
    > "$VERIFY_WORK/020-obt-driver.log" 2>&1
DRIVER_EXIT=$?
say "driver exit $DRIVER_EXIT, log $VERIFY_WORK/020-obt-driver.log"

# The package raises rather than answering a string, so a non-zero exit is the
# fault itself and its message is the diagnosis.
if [ "$DRIVER_EXIT" -ne 0 ]; then
    say "the last lines of the driver log:"
    tail -12 "$VERIFY_WORK/020-obt-driver.log" | sed 's/^/      /'
fi
assert_equal "$DRIVER_EXIT" "0" "the obt recipe committed without raising"

# WHAT LANDED. obt writes into <work>/linelist/.
LINELIST="$(sed -n 's/^OBT_VERIFY_LINELIST=//p' "$VERIFY_WORK/020-obt-driver.log" | head -1)"
if [ -z "$LINELIST" ]; then
    LINELIST="$(ls "$WORK"/linelist/*.xlsm "$WORK"/linelist/*.xlsb 2>/dev/null | head -1)"
fi

assert_file "$LINELIST" "obt delivered a linelist"
if [ -n "$LINELIST" ] && [ -f "$LINELIST" ]; then
    say "delivered: $LINELIST"
    assert_file_bigger_than "$LINELIST" 1000000 "the delivered linelist is a real workbook"

    # The same judge case 010 uses, so "coherent" means one thing in this suite
    # whichever path built the file.
    inspect_linelist "$LINELIST" "$VERIFY_WORK/020-obt-inspect.log"
    assert $? "linelist-inspect.R reports the obt-built linelist clean"
    if [ -s "${VERIFY_LAST_INSPECT_LOG:-}" ]; then
        grep -iE "fail|missing|absent" "${VERIFY_LAST_INSPECT_LOG}" | head -6 | sed 's/^/      /'
    fi

    # THE EXTENSION IS THE INTERESTING PART after the move to .xlsm. The
    # release binaries decide it, not src/: a channel still carrying the old
    # designer writes .xlsb, and one carrying the current code writes .xlsm.
    # Either is a pass, and which one it was is worth printing.
    say "obt delivered a $(printf '%s' "${LINELIST##*.}") linelist"
fi

check_case "the generic setup builds through the obt package"

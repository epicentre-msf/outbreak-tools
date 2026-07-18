#!/usr/bin/env bash
#
# interface-guard.sh — flag VBA interfaces that earn nothing.
#
# Enforces project-rules.md §4.4: an interface IFoo.cls is justified ONLY when
#   (a) two or more production classes Implements IFoo (real polymorphism), or
#   (b) a test fake Implements IFoo (a test seam), or
#   (c) it breaks a compile cycle (a human judgement call — NOT auto-detectable).
# Anything else (exactly one production implementer, no test fake) is 1:1
# ceremony and gets flagged.
#
# The check is advisory. Because case (c) cannot be detected mechanically, a
# flag means "review this", not "this is wrong". Use --write-baseline to record
# the interfaces that are knowingly pending removal (or are justified by (c)) so
# the guard only shouts about genuinely NEW offenders.
#
# Matching is by EXACT token. This matters: ICrossTable is a prefix of
# ICrossTableFormula, and ILinelist a prefix of ILinelistSpecs, so substring
# matching would mis-count. Test fakes live in BOTH src/tests/stubs AND
# src/tests/classes/stubs (and a few elsewhere under src/tests), so we scan all
# of src/tests.
#
# Targets bash 3.2 (stock macOS): no associative arrays, no mapfile.
#
# Usage:
#   interface-guard.sh                 report + (if a baseline exists) fail on new offenders
#   interface-guard.sh --report        report only, never fail (exit 0)
#   interface-guard.sh --write-baseline  snapshot current flagged interfaces as the baseline
#   interface-guard.sh -h | --help

set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
cd "$REPO_ROOT"

CLASSES_DIR="src/classes"
TESTS_DIR="src/tests"
LINELIST_FILE="src/classes/linelist/Linelist.cls"
BASELINE_FILE="$SCRIPT_DIR/interface-guard-baseline.txt"

# The 11 keepers from the interface-slimming decision (project-rules.md §4.4).
# They are all test seams, so case (b) already keeps them; this allowlist is a
# belt-and-suspenders safety net so they can never be flagged for removal.
# (Session H8 considered ICustomTest as a 12th keeper but REJECTED it: the
# Rubberduck.AssertClass asserter is assigned into `As Object` variables, never
# into `As ICustomTest`, so ICustomTest is 1:1 ceremony and folds — see H12.)
KEEPERS="IDesignerTranslation IDropdownLists ILLdictionary ILinelist ILinelistSpecs ITranslationObject"

MODE="check"
case "${1:-}" in
    -h|--help)
        sed -n '2,40p' "${BASH_SOURCE[0]:-$0}" | sed 's/^# \{0,1\}//'
        exit 0
        ;;
    --report)        MODE="report" ;;
    --write-baseline) MODE="write-baseline" ;;
    "")              MODE="check" ;;
    *)
        echo "interface-guard: unknown argument '$1' (try --help)" >&2
        exit 2
        ;;
esac

# --- tallies of `Implements I<Name>`, one line per file occurrence ----------
# Production implementers: src/classes excluding stale/.
PROD_TALLY="$(grep -rhoE 'Implements[[:space:]]+I[A-Za-z0-9_]+' "$CLASSES_DIR" \
    --include='*.cls' --exclude-dir=stale 2>/dev/null \
    | sed -E 's/^Implements[[:space:]]+//' | sort | uniq -c)"

# Test fakes: anywhere under src/tests (covers both stub folders + strays).
TEST_TALLY="$(grep -rhoE 'Implements[[:space:]]+I[A-Za-z0-9_]+' "$TESTS_DIR" \
    --include='*.cls' 2>/dev/null \
    | sed -E 's/^Implements[[:space:]]+//' | sort | uniq -c)"

# count_for <name> <tally>  -> exact-token count (0 if absent)
count_for() {
    printf '%s\n' "$2" | awk -v n="$1" '$2==n {c=$1} END {print c+0}'
}

is_keeper() {
    case " $KEEPERS " in *" $1 "*) return 0 ;; *) return 1 ;; esac
}

# fan_in <name> -> number of live (non-stale) files mentioning the token,
# excluding the interface's own definition file. Word-bounded so ICrossTable
# does not match ICrossTableFormula.
fan_in() {
    grep -rlw "$1" src --include='*.cls' --include='*.bas' \
        --exclude-dir=stale 2>/dev/null \
        | grep -v "/$1.cls$" | wc -l | tr -d ' '
}

in_linelist() {
    if grep -qw "$1" "$LINELIST_FILE" 2>/dev/null; then echo "yes"; else echo "no"; fi
}

# --- classify every interface ----------------------------------------------
# Rows: "<name>\t<prod>\t<test>\t<fanin>\t<linelist>\t<status>\t<reason>"
INTERFACES="$(find "$CLASSES_DIR" -name 'I*.cls' -not -path '*/stale/*' 2>/dev/null \
    | sed -E 's#.*/##; s/\.cls$//' | grep -E '^I[A-Z]' | sort -u)"

ROWS=""
FLAGGED=""
n_total=0 ; n_poly=0 ; n_seam=0 ; n_keep=0 ; n_flag=0

for name in $INTERFACES; do
    n_total=$((n_total + 1))
    prod="$(count_for "$name" "$PROD_TALLY")"
    test="$(count_for "$name" "$TEST_TALLY")"

    if is_keeper "$name"; then
        status="KEEP" ; reason="allowlist (11 keepers)" ; n_keep=$((n_keep + 1))
    elif [ "$prod" -ge 2 ]; then
        status="KEEP" ; reason="polymorphism ($prod impls)" ; n_poly=$((n_poly + 1))
    elif [ "$test" -ge 1 ]; then
        status="KEEP" ; reason="test seam ($test fake)" ; n_seam=$((n_seam + 1))
    else
        status="FLAG" ; reason="1:1 ceremony" ; n_flag=$((n_flag + 1))
        FLAGGED="$FLAGGED$name
"
    fi

    fanin="$(fan_in "$name")"
    ll="$(in_linelist "$name")"
    ROWS="$ROWS$name	$prod	$test	$fanin	$ll	$status	$reason
"
done

FLAGGED="$(printf '%s' "$FLAGGED" | grep -v '^$' || true)"

# --- write-baseline mode ----------------------------------------------------
if [ "$MODE" = "write-baseline" ]; then
    {
        echo "# interface-guard baseline — flagged interfaces known/accepted at snapshot time."
        echo "# Regenerate with: bash scripts/devtools/interface-guard.sh --write-baseline"
        echo "# The guard fails only on FLAGGED interfaces NOT listed here (genuinely new ones)."
        printf '%s\n' "$FLAGGED"
    } > "$BASELINE_FILE"
    echo "Wrote baseline ($(printf '%s\n' "$FLAGGED" | grep -c . ) flagged interfaces) -> $BASELINE_FILE"
    exit 0
fi

# --- report -----------------------------------------------------------------
echo "Interface guard (project-rules.md §4.4)"
echo "  scanned: $n_total interfaces in $CLASSES_DIR (excluding stale/)"
echo "  keep:    $n_keep allowlist + $n_poly polymorphic + $n_seam test-seam"
echo "  flag:    $n_flag 1:1 ceremony (unjustified)"
echo ""
printf '%-34s %4s %4s %5s %8s  %s\n' "INTERFACE" "PROD" "TEST" "FANIN" "LINELIST" "STATUS / REASON"
printf '%-34s %4s %4s %5s %8s  %s\n' "----------------------------------" "----" "----" "-----" "--------" "---------------"

# Flagged first, sorted easiest-first (fewest mentions); then the keepers.
printf '%s' "$ROWS" | grep -v '^$' | awk -F'\t' '$6=="FLAG"' | sort -t'	' -k4,4n \
    | while IFS='	' read -r name prod test fanin ll status reason; do
        printf '%-34s %4s %4s %5s %8s  %s (%s)\n' "$name" "$prod" "$test" "$fanin" "$ll" "$status" "$reason"
    done
printf '%s' "$ROWS" | grep -v '^$' | awk -F'\t' '$6=="KEEP"' | sort -t'	' -k1,1 \
    | while IFS='	' read -r name prod test fanin ll status reason; do
        printf '%-34s %4s %4s %5s %8s  %s (%s)\n' "$name" "$prod" "$test" "$fanin" "$ll" "$status" "$reason"
    done
echo ""

# --- exit policy ------------------------------------------------------------
if [ "$MODE" = "report" ]; then
    exit 0
fi

if [ ! -f "$BASELINE_FILE" ]; then
    echo "No baseline found. Guard is not armed."
    echo "  Run: bash scripts/devtools/interface-guard.sh --write-baseline"
    echo "  to record the $n_flag currently-flagged interfaces as accepted; thereafter"
    echo "  the guard fails only on NEW unjustified interfaces."
    exit 0
fi

# Compare flagged set against the baseline; fail only on genuinely new offenders.
BASE_NAMES="$(grep -v '^#' "$BASELINE_FILE" 2>/dev/null | grep -v '^$' | sort -u)"
NEW_OFFENDERS=""
for name in $FLAGGED; do
    case "
$BASE_NAMES
" in
        *"
$name
"*) : ;;
        *) NEW_OFFENDERS="$NEW_OFFENDERS$name
" ;;
    esac
done
NEW_OFFENDERS="$(printf '%s' "$NEW_OFFENDERS" | grep -v '^$' || true)"

if [ -n "$NEW_OFFENDERS" ]; then
    echo "FAIL: new unjustified interface(s) added since the baseline:"
    printf '%s\n' "$NEW_OFFENDERS" | sed 's/^/  - /'
    echo ""
    echo "Each must satisfy project-rules.md §4.4 case (a) two impls, (b) a test fake,"
    echo "or (c) a compile-cycle break. If (c) applies, record it with --write-baseline."
    exit 1
fi

echo "OK: no new unjustified interfaces beyond the baseline."
exit 0

#!/usr/bin/env bash
#
# show-names.sh — list the class or module names that live in one source folder.
#
# The sources are split by kind and then by folder:
#
#   src/classes/<folder>/*.cls     the classes
#   src/modules/<folder>/*.bas     the modules
#   src/tests/<folder>/*.cls|*.bas the test suites and their fixtures
#
# This prints the names in one of those folders, one per line, sorted. Names
# come out bare by default, which is the form the test registry and the
# `@depends` lines want. Ask for --ext when you need the file names instead.
#
# Targets bash 3.2 (stock macOS): no associative arrays, no mapfile.
#
# Usage:
#   show-names.sh --classes --folder=dataio      the 7 class names in src/classes/dataio
#   show-names.sh --classes --folder=dataio --ext        the same, as file names
#   show-names.sh --modules --folder=linelist    the module names
#   show-names.sh --tests   --folder=dataio      the test files
#   show-names.sh --classes                      every class, all folders
#   show-names.sh --classes --folders            the folder names under src/classes
#   show-names.sh --classes --folder=dataio --paths      paths instead of names
#   show-names.sh -h | --help

set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
cd "$REPO_ROOT"

usage() {
    sed -n '2,30p' "${BASH_SOURCE[0]:-$0}" | sed 's/^# \{0,1\}//'
}

KIND=""
FOLDER=""
WANT_EXT="no"
WANT_PATHS="no"
LIST_FOLDERS="no"

while [ $# -gt 0 ]; do
    case "$1" in
        -h|--help)   usage ; exit 0 ;;
        --classes)   KIND="classes" ;;
        --modules)   KIND="modules" ;;
        --tests)     KIND="tests" ;;
        --ext)       WANT_EXT="yes" ;;
        --paths)     WANT_PATHS="yes" ;;
        --folders)   LIST_FOLDERS="yes" ;;
        --folder=*)  FOLDER="${1#--folder=}" ;;
        --folder)    shift ; FOLDER="${1:-}" ;;
        *)
            echo "show-names: unknown argument '$1' (try --help)" >&2
            exit 2
            ;;
    esac
    shift
done

if [ -z "$KIND" ]; then
    echo "show-names: pick one of --classes, --modules, --tests (try --help)" >&2
    exit 2
fi

ROOT="src/$KIND"
if [ ! -d "$ROOT" ]; then
    echo "show-names: no such folder '$ROOT'" >&2
    exit 1
fi

# --folders: just the folder names under this kind.
if [ "$LIST_FOLDERS" = "yes" ]; then
    find "$ROOT" -mindepth 1 -maxdepth 1 -type d | sed -E 's#.*/##' | sort
    exit 0
fi

# Where to look. An empty --folder means every folder of this kind.
if [ -n "$FOLDER" ]; then
    DIR="$ROOT/$FOLDER"
    if [ ! -d "$DIR" ]; then
        echo "show-names: no such folder '$DIR'. Folders here are:" >&2
        find "$ROOT" -mindepth 1 -maxdepth 1 -type d | sed -E 's#.*/##' | sort | sed 's/^/  /' >&2
        exit 1
    fi
else
    DIR="$ROOT"
fi

# Classes are .cls, modules are .bas, tests hold both.
case "$KIND" in
    classes) FILES="$(find "$DIR" -name '*.cls' -type f | sort)" ;;
    modules) FILES="$(find "$DIR" -name '*.bas' -type f | sort)" ;;
    tests)   FILES="$(find "$DIR" \( -name '*.cls' -o -name '*.bas' \) -type f | sort)" ;;
esac

if [ -z "$FILES" ]; then
    exit 0
fi

if [ "$WANT_PATHS" = "yes" ]; then
    printf '%s\n' "$FILES"
elif [ "$WANT_EXT" = "yes" ]; then
    printf '%s\n' "$FILES" | sed -E 's#.*/##' | sort
else
    printf '%s\n' "$FILES" | sed -E 's#.*/##; s/\.(cls|bas)$//' | sort
fi

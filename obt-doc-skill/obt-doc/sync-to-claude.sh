#!/bin/bash
# Install or sync OBT-Doc skill to Claude's skill directory
# Run from anywhere -- the script detects its own location.
#
# First run  : creates ~/.claude/skills/obt-doc and copies all files.
# Later runs : overwrites only the files that belong to the skill.
#
# Usage:
#   ./sync-to-claude.sh            # install / update
#   ./sync-to-claude.sh --check    # dry-run: show what would be copied

set -e

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SOURCE_DIR="$SCRIPT_DIR"
TARGET_DIR="$HOME/.claude/skills/obt-doc"

DRY_RUN=false
if [ "${1:-}" = "--check" ]; then
    DRY_RUN=true
fi

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
copy_file() {
    local src="$1"
    local dst="$2"

    if [ "$DRY_RUN" = true ]; then
        echo "  [dry-run] $src -> $dst"
    else
        cp "$src" "$dst"
        echo "  copied  $src -> $dst"
    fi
}

# ---------------------------------------------------------------------------
# Banner
# ---------------------------------------------------------------------------
echo ""
echo "  OBT-Doc Skill Installer / Sync"
echo "  ================================"
echo "  Source : $SOURCE_DIR"
echo "  Target : $TARGET_DIR"
echo ""

# ---------------------------------------------------------------------------
# Create target directory tree on first run
# ---------------------------------------------------------------------------
if [ ! -d "$TARGET_DIR" ]; then
    if [ "$DRY_RUN" = true ]; then
        echo "  [dry-run] Would create $TARGET_DIR"
        echo "  [dry-run] Would create $TARGET_DIR/references"
    else
        echo "  First install -- creating target directory..."
        mkdir -p "$TARGET_DIR/references"
    fi
fi

# Ensure references/ subdirectory exists (may be missing after manual edits)
if [ "$DRY_RUN" = false ] && [ ! -d "$TARGET_DIR/references" ]; then
    mkdir -p "$TARGET_DIR/references"
fi

# ---------------------------------------------------------------------------
# Sync core files
# ---------------------------------------------------------------------------
echo ""
echo "  Syncing core files..."
copy_file "$SOURCE_DIR/SKILL.md" "$TARGET_DIR/SKILL.md"

# ---------------------------------------------------------------------------
# Sync reference files
# ---------------------------------------------------------------------------
echo ""
echo "  Syncing reference files..."
for ref in "$SOURCE_DIR"/references/*.md; do
    [ -f "$ref" ] || continue
    copy_file "$ref" "$TARGET_DIR/references/$(basename "$ref")"
done

# ---------------------------------------------------------------------------
# Sync optional top-level files (README, skill.json, etc.)
# ---------------------------------------------------------------------------
echo ""
echo "  Syncing optional files..."
for optional in README.md skill.json; do
    if [ -f "$SOURCE_DIR/$optional" ]; then
        copy_file "$SOURCE_DIR/$optional" "$TARGET_DIR/$optional"
    fi
done

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
echo ""
if [ "$DRY_RUN" = true ]; then
    echo "  Dry run complete -- no files were modified."
else
    echo "  Sync complete!"
    echo ""
    echo "  Next steps:"
    echo "    1. Reload your editor window (Cmd/Ctrl+Shift+P -> 'Reload Window')"
    echo "    2. Or restart Claude to pick up the updated skill"
fi
echo ""

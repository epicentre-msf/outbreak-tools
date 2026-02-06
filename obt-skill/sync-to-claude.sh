#!/bin/bash
# Sync OBT skill changes from source to installed location
# Run this after modifying skill files in obt-skill/

set -e

# Detect source directory from script location
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SOURCE_DIR="$SCRIPT_DIR"
TARGET_DIR="$HOME/.claude/skills/obt"

echo "🔄 Syncing OBT skill..."
echo "  From: $SOURCE_DIR"
echo "  To:   $TARGET_DIR"
echo ""

# Check if target exists
if [ ! -d "$TARGET_DIR" ]; then
    echo "❌ Target directory doesn't exist: $TARGET_DIR"
    echo "Run install-vscode.sh first to install the skill."
    exit 1
fi

# Sync skill files
echo "📋 Copying skill files..."
cp "$SOURCE_DIR/SKILL.md" "$TARGET_DIR/"
cp "$SOURCE_DIR/project-rules.md" "$TARGET_DIR/"
cp "$SOURCE_DIR/skill.json" "$TARGET_DIR/"
cp "$SOURCE_DIR/README.md" "$TARGET_DIR/"

# Copy optional files if they exist
[ -f "$SOURCE_DIR/ARCHITECTURE-OVERVIEW.md" ] && cp "$SOURCE_DIR/ARCHITECTURE-OVERVIEW.md" "$TARGET_DIR/"
[ -f "$SOURCE_DIR/MIGRATION-GUIDE.md" ] && cp "$SOURCE_DIR/MIGRATION-GUIDE.md" "$TARGET_DIR/"
[ -f "$SOURCE_DIR/UPDATES.md" ] && cp "$SOURCE_DIR/UPDATES.md" "$TARGET_DIR/"

echo ""
echo "✅ Sync complete!"
echo ""
echo "📋 Next steps:"
echo "  1. Reload VSCode window (Cmd/Ctrl+Shift+P → 'Reload Window')"
echo "  2. Test: /obt help"
echo ""

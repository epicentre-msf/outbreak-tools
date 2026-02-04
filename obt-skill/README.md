# OutbreakTools VBA Skill (OBT)

**Version:** 1.0
**Created:** 2026-02-04
**Purpose:** Specialized VBA development assistant for the OutbreakTools project

---

## 📁 Architecture Overview

### Directory Structure

```
outbreak-tools/
├── obt-skill/                    ← Skill package (this directory)
│   ├── SKILL.md                  ← Main skill definition
│   ├── project-rules.md          ← Comprehensive coding standards
│   └── README.md                 ← This file
│
├── .obt/                         ← Active workspace (NEW)
│   ├── implementations.md        ← Current task specifications
│   └── tracking.md               ← Progress tracking
│
├── instructions.md               ← LEGACY (reference only)
├── tracking.md                   ← LEGACY (reference only)
└── implementations.md            ← LEGACY (reference only)
```

### Legacy vs Active Files

| File Type | Legacy Location | Active Location | Status |
|-----------|----------------|-----------------|---------|
| Instructions | `instructions.md` | `obt-skill/project-rules.md` | Legacy = Reference |
| Tracking | `tracking.md` | `.obt/tracking.md` | Legacy = Archived |
| Specifications | `implementations.md` | `.obt/implementations.md` | Legacy = Archived |

**Important:** The skill references the NEW `.obt/` directory. Legacy files remain untouched as historical reference.

---

## 🚀 Installation (For User)

### Option 1: Manual Installation (Recommended for Review)

1. **Review the skill files in `obt-skill/` directory:**
   - Read `SKILL.md` to understand workflow
   - Read `project-rules.md` to verify coding standards
   - Ensure all rules from legacy `instructions.md` are properly captured

2. **Copy skill to Claude's skill directory:**
   ```bash
   cp -r obt-skill ~/.claude/skills/obt
   # Or wherever your Claude skills are installed
   ```

3. **Restart Claude or reload skills:**
   ```bash
   claude skills reload
   # Or restart Claude application
   ```

### Option 2: Skill Package Installation

If packaging as a plugin:
1. Package the `obt-skill/` directory as a skill bundle
2. Install via Claude's skill management interface
3. Verify installation with `/skills list`

---

## 📖 How It Works

### Automatic Triggering

The skill automatically activates when you use VBA-related terms:

**Examples that trigger the skill:**
- "Refactor the ReportBuilder class"
- "Create a new VBA module for data validation"
- "Fix the bug in src/classes/LLFormat.cls"
- "Add tests for the new DataExporter class"

### Skill Workflow

When triggered, the OBT skill follows this process:

```
1. READ CONTEXT
   ├─ obt-skill/project-rules.md    (coding standards)
   ├─ .obt/implementations.md        (current task spec)
   └─ .obt/tracking.md               (progress state)

2. ANALYZE REQUEST
   ├─ Which files need modification?
   ├─ Does this align with implementations.md?
   └─ Check constraints in project-rules.md

3. IMPLEMENT CHANGES
   ├─ Follow all coding standards
   ├─ Return FULL file contents (never diffs)
   └─ Create tests for new classes

4. POST-PROCESS
   ├─ Run unix2dos on modified files
   ├─ Update .obt/tracking.md with progress
   └─ Document next steps if incomplete
```

### Manual Invocation

You can also explicitly invoke the skill:

```bash
/obt refactor the ReportBuilder class to use ProjectError pattern
```

---

## 📝 Using the Workflow Files

### `.obt/implementations.md`

**Purpose:** Define what you want to accomplish
**Updated by:** User (manually)
**Read by:** AI assistant

**When to update:**
- Starting a new major task
- Changing project goals
- Adding new requirements

**Example:**
```markdown
Goal:
Add export functionality to LLFormat class

Behaviors:
- Export takes a workbook parameter
- Skip export if worksheet exists, import instead
- Create worksheet from scratch with all required elements

Rules:
- Follow project-rules.md
- Use ProjectError for error handling
```

### `.obt/tracking.md`

**Purpose:** Track implementation progress
**Updated by:** AI assistant (automatically)
**Read by:** Both user and AI

**AI updates this file:**
- Marking tasks as [DONE]
- Updating State section after each work session
- Documenting stopping points if work is incomplete

**You monitor it to:**
- See progress on current task
- Understand what's been completed
- Know what's next

---

## 🎯 Key Features

### ✅ Enforced Standards

The skill enforces OutbreakTools coding standards:
- Full file rewrites (never diffs/snippets)
- Naming conventions (camelCase variables, PascalCase classes)
- Error handling with ProjectError pattern
- Cross-platform compatibility (Windows & macOS)
- No external DLLs or Dictionary objects
- unix2dos post-processing

### ✅ Architectural Compliance

- Respects `src/classes/` and `src/tests/` structure
- Adds tests for new classes automatically
- Prevents circular dependencies
- Never modifies protected files (e.g., DictionaryTestFixture.bas)

### ✅ Workflow Integration

- Reads task specifications from `.obt/implementations.md`
- Tracks progress in `.obt/tracking.md`
- Documents stopping points for seamless resumption
- References comprehensive rules in `project-rules.md`

---

## 🔍 Example Usage

### Example 1: Refactoring a Class

**User request:**
```
Refactor the ReportBuilder class to improve error handling using ProjectError
```

**Skill workflow:**
1. Reads `project-rules.md` for error handling standards
2. Reads `.obt/implementations.md` to understand current goals
3. Reads `.obt/tracking.md` to see if this is part of ongoing work
4. Reads `src/classes/ReportBuilder.cls`
5. Refactors code following all standards
6. Returns FULL refactored file
7. Runs `unix2dos src/classes/ReportBuilder.cls`
8. Updates `.obt/tracking.md` with progress

### Example 2: Creating a New Class

**User request:**
```
Create a new DataValidator class with methods to validate table structures
```

**Skill workflow:**
1. Reads `project-rules.md` for class creation standards
2. Creates `src/classes/DataValidator.cls` with interface
3. Creates `src/tests/DataValidatorTestFixture.bas`
4. Follows all naming and documentation conventions
5. Returns both complete files
6. Runs unix2dos on both files
7. Updates `.obt/tracking.md` with new files created

---

## 🛠️ Customization

### Modifying Coding Standards

To update coding standards:
1. Edit `obt-skill/project-rules.md`
2. The skill will use updated rules immediately (no reinstall needed)

### Adding New Rules

To add project-specific rules:
1. Add rules to appropriate section in `project-rules.md`
2. Consider adding to Section 12 (Specific Coding Rules) for quick reference

### Changing Workflow

To modify the workflow:
1. Edit `obt-skill/SKILL.md`
2. Update Phase sections as needed
3. Ensure workflow still reads `project-rules.md` first

---

## 📋 Verification Checklist

Before using the skill, verify:

- [ ] `SKILL.md` contains correct trigger keywords
- [ ] `project-rules.md` has all standards from legacy `instructions.md`
- [ ] `.obt/implementations.md` exists and is writable
- [ ] `.obt/tracking.md` exists and is writable
- [ ] File paths in SKILL.md match your workspace structure
- [ ] Legacy files (`instructions.md`, `tracking.md`, `implementations.md`) remain untouched

---

## 🐛 Troubleshooting

### Skill Not Triggering

**Problem:** Skill doesn't activate with VBA keywords
**Solution:** Check trigger keywords in `SKILL.md` and verify skill is installed

### Wrong Files Referenced

**Problem:** Skill reads legacy tracking.md instead of `.obt/tracking.md`
**Solution:** Verify paths in `SKILL.md` Phase 1 section

### Incomplete File Returns

**Problem:** Skill returns partial file or diffs
**Solution:** Check that Section 2.1 in `project-rules.md` is being followed

---

## 📞 Support

For issues or questions:
1. Check `project-rules.md` for coding standards
2. Review `SKILL.md` for workflow steps
3. Verify `.obt/` files are properly structured
4. Consult legacy `instructions.md` for historical context

---

## 🔄 Version History

- **1.0** (2026-02-04): Initial skill creation
  - Extracted rules from legacy `instructions.md`
  - Created `.obt/` workspace architecture
  - Established automatic triggering and workflow

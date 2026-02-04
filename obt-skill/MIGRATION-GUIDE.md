# Migration Guide: Legacy to OBT Skill Architecture

**Date:** 2026-02-04
**Purpose:** Guide for transitioning from legacy instruction files to the OBT skill system

---

## 🔄 What Changed

### Before (Legacy System)

```
outbreak-tools/
├── instructions.md          ← Manual reference, easy to miss
├── tracking.md              ← Manual updates, inconsistent
└── implementations.md       ← Manual coordination required
```

**Problems:**
- ❌ No automatic triggering - Claude might forget to read instructions.md
- ❌ No systematic enforcement of coding standards
- ❌ Manual workflow coordination required
- ❌ Easy to skip tracking updates
- ❌ No validation that rules are followed

### After (OBT Skill System)

```
outbreak-tools/
├── obt-skill/
│   ├── SKILL.md             ← Automatic workflow orchestration
│   ├── project-rules.md     ← Comprehensive standards (from instructions.md)
│   └── README.md            ← Documentation
│
├── .obt/
│   ├── implementations.md   ← Active task specification
│   └── tracking.md          ← Active progress tracking
│
├── instructions.md          ← LEGACY (kept as reference)
├── tracking.md              ← LEGACY (archived work)
└── implementations.md       ← LEGACY (historical context)
```

**Improvements:**
- ✅ Automatic triggering on VBA keywords
- ✅ Systematic enforcement of all coding standards
- ✅ Guaranteed workflow: read rules → implement → track progress
- ✅ Automatic tracking updates
- ✅ Consistent, repeatable process

---

## 📋 Migration Steps

### Step 1: Review Skill Files ✓ (COMPLETE)

All skill files have been created in `obt-skill/` directory:
- [x] SKILL.md - Workflow and triggers
- [x] project-rules.md - Coding standards
- [x] README.md - Usage documentation
- [x] MIGRATION-GUIDE.md - This file

### Step 2: Review Standards Mapping

Verify that all rules from legacy `instructions.md` are captured in `project-rules.md`:

| Legacy Section | New Location | Status |
|----------------|--------------|--------|
| §1: Scope and Role | project-rules.md §1 | ✅ Mapped |
| §2: File Handling | project-rules.md §2 | ✅ Mapped |
| §3: VBA Constraints | project-rules.md §3 | ✅ Mapped |
| §4: Architecture | project-rules.md §4 | ✅ Mapped |
| §5: Naming | project-rules.md §5 | ✅ Mapped |
| §6: Documentation | project-rules.md §6 | ✅ Mapped |
| §7: Error Handling | project-rules.md §7 | ✅ Mapped |
| §8: Style | project-rules.md §8 | ✅ Mapped |
| §9: Assumptions | project-rules.md §10 | ✅ Mapped |
| §10: Output | project-rules.md §11 | ✅ Mapped |
| §11: Post-processing | project-rules.md §12 | ✅ Mapped |
| §12: Specific Rules | project-rules.md §13 | ✅ Mapped |
| §13: When in Doubt | project-rules.md §14 | ✅ Mapped |

**Action:** Review `obt-skill/project-rules.md` and compare with `instructions.md` to ensure nothing was missed.

### Step 3: Migrate Active Work (If Any)

If you have active work tracked in legacy files:

**A. Migrate implementations.md:**
```bash
# Copy current goal/behaviors from legacy implementations.md
# Paste into .obt/implementations.md
# Update format to match new template
```

**B. Migrate tracking.md:**
```bash
# Copy current tasks and state from legacy tracking.md
# Paste into .obt/tracking.md
# Update format to match new template
```

**Example migration:**

**From** `tracking.md` (legacy):
```markdown
Goal
- Add LLFormat export workflow

Tasks
- [DONE] Review LLFormat/ILLFormat structure
- [DONE] Add Export member to interface
- [DONE] Build export logic

State
- Latest update (2026-01-21): LLFormat export implemented
```

**To** `.obt/tracking.md` (active):
```markdown
Goal
Add LLFormat export workflow

Tasks
- [DONE] Review LLFormat/ILLFormat structure
- [DONE] Add Export member to interface
- [DONE] Build export logic
- Verify export works with existing setup code

State
Latest update (2026-02-04):
Migrated tracking from legacy file. LLFormat export is complete.
Next: Test integration with SetupImportService.
```

### Step 4: Install the Skill

**Option A: Local Installation (if Claude skills directory is accessible)**
```bash
# Copy skill directory to Claude skills location
cp -r obt-skill ~/.claude/skills/obt

# Reload skills
claude skills reload
```

**Option B: Package Installation (if using plugin system)**
```bash
# Package as skill bundle
# Install via Claude's skill management
# Follow platform-specific installation process
```

**Verify installation:**
```bash
# Check if skill is listed
claude skills list | grep obt

# Or test by asking Claude:
"List available skills"
```

### Step 5: Test the Skill

**Test 1: Automatic Triggering**
```
User: "Review the ReportBuilder class"
Expected: OBT skill should automatically activate
```

**Test 2: Workflow Execution**
```
User: "Fix error handling in DataExporter.cls"
Expected:
1. Skill reads project-rules.md
2. Skill reads .obt/implementations.md
3. Skill reads .obt/tracking.md
4. Skill implements changes
5. Skill updates .obt/tracking.md
```

**Test 3: Standards Enforcement**
```
User: "Create a new VBA class for validation"
Expected:
1. Class created in src/classes/
2. Test created in src/tests/
3. Full file returned (not snippet)
4. unix2dos executed
5. Tracking updated
```

### Step 6: Archive Legacy Files (Optional)

Once the skill is working and tested:

**Option A: Keep as Reference (Recommended)**
```bash
# Leave legacy files in place
# They serve as historical reference
# No action needed
```

**Option B: Move to Archive Directory**
```bash
mkdir archive-legacy-instructions
mv instructions.md archive-legacy-instructions/
mv tracking.md archive-legacy-instructions/
mv implementations.md archive-legacy-instructions/
```

**Option C: Add Git Tag**
```bash
git add -A
git commit -m "Migrate to OBT skill architecture

- Created obt-skill/ with SKILL.md and project-rules.md
- Established .obt/ workspace for active tracking
- Legacy instruction files preserved for reference"
git tag -a v1.0-obt-skill -m "OBT skill architecture migration"
```

---

## 🎯 How to Use the New System

### Starting a New Task

**1. Define the task in `.obt/implementations.md`:**
```markdown
Goal:
Add data validation to report generation

Behaviors:
- Validate source data before processing
- Raise clear errors for invalid data
- Log validation failures

Rules:
- Use ProjectError for errors
- Add tests for each validation rule
```

**2. Initialize tracking in `.obt/tracking.md`:**
```markdown
Goal
Add data validation to report generation

Tasks
- Create DataValidator class
- Implement validation rules
- Add error handling
- Write tests
- Integrate with ReportBuilder

State
Latest update (2026-02-04):
Task initialized, ready to start implementation.
```

**3. Invoke the skill:**
```
"Create the DataValidator class with validation for report data"
```

**4. Skill executes automatically:**
- Reads project-rules.md
- Reads implementations.md (your spec)
- Reads tracking.md (current state)
- Implements the class
- Updates tracking.md with progress

### Continuing Existing Work

**1. Review tracking:**
```bash
cat .obt/tracking.md
# See what was done and what's next
```

**2. Continue with next task:**
```
"Implement the next validation rule in DataValidator"
```

**3. Skill resumes work:**
- Sees [DONE] markers in tracking.md
- Knows what's been completed
- Works on next pending task
- Updates tracking automatically

---

## 📊 Comparison: Legacy vs New

| Aspect | Legacy System | OBT Skill System |
|--------|---------------|------------------|
| **Triggering** | Manual reminders | Automatic on VBA keywords |
| **Standards** | Manual reference | Enforced automatically |
| **Tracking** | Inconsistent updates | Automatic updates |
| **Workflow** | Ad-hoc | Systematic (4 phases) |
| **File Returns** | Sometimes partial | Always full files |
| **unix2dos** | Often forgotten | Always executed |
| **Testing** | Sometimes skipped | Always added for new classes |
| **Documentation** | Sometimes incomplete | Tags always preserved |
| **Reliability** | Depends on memory | Guaranteed process |

---

## ⚠️ Important Notes

### Legacy Files Are NOT Deleted

The legacy `instructions.md`, `tracking.md`, and `implementations.md` are:
- ✅ Preserved as reference
- ✅ Available for historical context
- ✅ Useful for comparing old vs new work
- ❌ No longer actively used by the skill

### The Skill References `.obt/` Directory

All active work uses:
- `.obt/implementations.md` (not root implementations.md)
- `.obt/tracking.md` (not root tracking.md)
- `obt-skill/project-rules.md` (not root instructions.md)

### Both Systems Can Coexist

If needed, you can:
- Use OBT skill for new VBA work
- Reference legacy files for historical context
- Keep both systems in parallel during transition

---

## ✅ Post-Migration Checklist

After migration, verify:

- [ ] OBT skill is installed and appears in skills list
- [ ] Skill triggers automatically on VBA keywords
- [ ] Skill reads `obt-skill/project-rules.md` first
- [ ] Skill reads `.obt/implementations.md` for task specs
- [ ] Skill reads `.obt/tracking.md` for progress
- [ ] Skill updates `.obt/tracking.md` after work
- [ ] Full files are returned (never diffs/snippets)
- [ ] unix2dos runs automatically
- [ ] Tests are created for new classes
- [ ] Legacy files remain untouched

---

## 🚀 Next Steps

1. **Review** all files in `obt-skill/` directory
2. **Verify** project-rules.md captures all standards
3. **Migrate** any active work to `.obt/` files
4. **Install** the skill in Claude
5. **Test** with a simple VBA task
6. **Monitor** `.obt/tracking.md` for automatic updates
7. **Enjoy** consistent, high-quality VBA development!

---

## 📞 Questions?

If you have questions about:
- **Standards:** Check `obt-skill/project-rules.md`
- **Workflow:** Check `obt-skill/SKILL.md`
- **Usage:** Check `obt-skill/README.md`
- **Migration:** This file
- **Legacy context:** Original `instructions.md`

Happy coding! 🎉

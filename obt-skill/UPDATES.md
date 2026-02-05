# OBT Skill Updates - 2026-02-04

## Changes Based on User Feedback

The following updates were made to reflect the actual project structure and workflow requirements:

---

## 1. ✅ Subfolder Structure Added

### Classes Folder Organization
**Before:** `src/classes/*.cls` (flat structure assumed)
**After:** `src/classes/<topic>/*.cls` (subfolders by topic)

Classes are now organized in subfolders by topic/feature area, with each subfolder containing related classes.

### Modules Folder Added
**New structure:** `src/modules/<topic>/*.bas`

Modules (orchestration .bas files) are organized in subfolders that mirror the classes structure. Each topic has corresponding modules for orchestration.

---

## 2. ✅ Stale Mode Checking (CRITICAL)

### New Workflow Step
Added **mandatory stale-mode checking** before working on any file in a `stale/` subfolder:

**Location in SKILL.md:** Phase 2, Step 2.2
**Location in project-rules.md:** Section 4.2

**Process:**
1. AI detects file is in `stale/` folder
2. AI **STOPS immediately**
3. AI **ASKS user:** "The file [filename] is in the stale folder. Is this class/module ready to use, or is it in stale mode (needs deletion/refactoring)?"
4. AI **WAITS** for explicit user confirmation
5. Only proceeds if user confirms

**Why:** Stale folders contain files that may:
- Need deletion
- Require major refactoring
- Be incomplete/non-functional
- Be outdated

---

## 3. ✅ Legacy Folders Protection

### New Rules Added
**Location in SKILL.md:** Phase 2, Step 2.3
**Location in project-rules.md:** Section 4.1 & 4.3

**Rules:**
- ❌ **NEVER modify files in legacy folders** within `src/`
- ❌ Never delete legacy folders or their contents
- ❌ Never move files into or out of legacy folders
- ✅ Only user manages legacy folders

**Why:** Legacy folders will be progressively removed by the user. AI should not interfere with this cleanup process.

---

## 4. ✅ File Creation Workflow Enhanced

### Ask User for Topic Subfolder
**Location in SKILL.md:** Phase 3, Step 3.3
**Location in project-rules.md:** Section 4.3

**New workflow when creating files:**

**For Classes:**
1. AI asks: "Which topic subfolder should I create this class in? (e.g., reporting, validation, etc.)"
2. User specifies folder
3. AI creates in `src/classes/<topic>/ClassName.cls`

**For Modules:**
1. AI asks: "Which topic subfolder should I create this module in?"
2. User specifies folder
3. AI creates in `src/modules/<topic>/ModuleName.bas`

---

## 5. ✅ Updated Documentation

### Files Updated

**SKILL.md:**
- Phase 1, Step 1.3: Updated workspace structure diagram
- Phase 2, Step 2.2: Added stale-mode checking
- Phase 2, Step 2.3: Added legacy folder protection
- Phase 3, Step 3.3: Added subfolder selection for new files
- Quick Reference: Updated file locations

**project-rules.md:**
- Section 4.1: Complete project structure with subfolders
- Section 4.2: New section on stale-mode checking
- Section 4.3: Updated architectural principles
- Section 4.3: Added module design guidance
- Section 15: Updated file locations quick reference

---

## 6. ✅ Error Handling Updates (User-Added)

The user has already updated `project-rules.md` with:
- Enhanced error handling pattern (Section 7.1)
- `ThrowError` helper sub pattern
- Updated test structure using `CustomTest` (Section 9.1)
- Reference to `TestLLChoices.bas` for test inspiration

These changes were preserved and integrated into the skill.

---

## Complete Structure Now Documented

```
src/
├── classes/
│   ├── <topic1>/          ← Classes for specific topic
│   ├── <topic2>/          ← Classes for specific topic
│   ├── stale/             ← ⚠️ CHECK before use
│   └── ...
│
├── modules/
│   ├── <topic1>/          ← Modules for specific topic
│   ├── <topic2>/          ← Modules for specific topic
│   ├── stale/             ← ⚠️ CHECK before use
│   └── ...
│
├── tests/
│   ├── *TestFixture.bas
│   └── TestHelpers.bas
│
└── [legacy folders]/      ← ❌ DO NOT TOUCH
```

---

## Workflow Changes Summary

### Before Updates
- AI might work on stale files without checking
- No module folder documented
- Flat file structure assumed
- No protection for legacy folders
- No guidance on topic organization

### After Updates
- ✅ Mandatory stale-mode checking with user confirmation
- ✅ Module folder fully documented
- ✅ Subfolder structure by topic documented
- ✅ Legacy folders protected from AI modifications
- ✅ AI asks user which topic subfolder to use for new files

---

## Verification

All changes maintain:
- ✅ Full file return policy (never diffs)
- ✅ unix2dos post-processing
- ✅ Tracking updates
- ✅ Test creation for new classes
- ✅ All original coding standards
- ✅ User's error handling enhancements

---

## Ready for Review

The skill now accurately reflects:
1. **Actual folder structure** with topics and stale folders
2. **Modules folder** for orchestration code
3. **Safety checks** for stale files
4. **Protection** for legacy folders
5. **User guidance** for topic selection

**Status:** ✅ All user feedback incorporated
**Next:** Ready for packaging after final review

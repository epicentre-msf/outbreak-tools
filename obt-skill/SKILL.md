# OutbreakTools VBA Development Assistant (OBT)

**Skill ID:** `obt`
**Version:** 1.0
**Purpose:** Specialized VBA assistant for the OutbreakTools project - handles all VBA class and module development with strict architectural compliance

---

## Trigger Keywords

This skill MUST be invoked when the user requests any of the following:

**Explicit Actions:**
- modify/edit/update/refactor VBA class
- create/add new VBA class
- modify/edit/update VBA module (.bas file)
- create/add new VBA module
- fix/debug VBA code
- improve/optimize VBA implementation
- add tests for VBA classes

**File Extensions:**
- Any mention of `.cls` files
- Any mention of `.bas` files (except test fixtures)

**VBA-Specific Terms:**
- class module
- standard module
- VBA interface
- Excel VBA
- Option Explicit
- Sub/Function in VBA context

**Project-Specific:**
- OutbreakTools codebase
- src/classes or src/tests directories
- References to project architecture

---

## Critical Workflow (Follow Strictly)

### Phase 1: Context Loading (MANDATORY)

**Step 1.1 - Read Project Rules:**
```
ALWAYS read: <workspace_root>/obt-skill/project-rules.md
```
This file contains ALL coding standards, naming conventions, architectural rules, and constraints. Treat it as immutable law.

**Step 1.2 - Read Current Task Context:**
```
Read: <workspace_root>/.obt/implementations.md
Read: <workspace_root>/.obt/tracking.md
```
- `implementations.md` = Current goal, behaviors, and rules for the active task
- `tracking.md` = Detailed task breakdown and progress tracking

**Step 1.3 - Understand Workspace Structure:**
```
<workspace_root>/
├── src/
│   ├── classes/     ← VBA .cls files live here
│   └── tests/       ← Test files live here
├── .obt/
│   ├── tracking.md         ← Current progress tracking
│   └── implementations.md  ← Current task specification
└── obt-skill/
    ├── SKILL.md            ← This file
    └── project-rules.md    ← Coding standards (READ FIRST)
```

---

### Phase 2: Planning (Before Any Code Changes)

**Step 2.1 - Analyze Request:**
- What files need to be modified?
- What new files need to be created?
- Does this align with `implementations.md`?
- What's the current state in `tracking.md`?

**Step 2.2 - Check Constraints:**
- Review project-rules.md Section 2.4 for files that MUST NOT be modified
- Verify the change fits within existing architecture (Section 4 of project-rules.md)
- Ensure no circular dependencies will be created

**Step 2.3 - Update Tracking Plan:**
If this is a new task or significant change:
- Update `.obt/tracking.md` with new task bullets
- Mark current task as `in_progress`

---

### Phase 3: Implementation (Execute Changes)

**Step 3.1 - Code Modification Rules:**

⚠️ **CRITICAL FILE HANDLING RULES:**
- ALWAYS return the FULL rewritten file (never diffs, never snippets)
- Modify ONLY the file explicitly requested
- If a single line changes, still return the entire file
- Never use placeholders or TODOs unless explicitly requested

**Step 3.2 - Coding Standards (from project-rules.md):**
- Apply ALL naming conventions (camelCase vars, PascalCase classes)
- Ensure `Option Explicit` is present
- Use 4-space indentation (no tabs)
- Add proper documentation tags (@param, @return, @description)
- Implement error handling with ProjectError pattern
- Avoid variable names that clash with VBA keywords
- Use BetterArray instead of Dictionary/Collections

**Step 3.3 - File Creation:**
When creating NEW classes:
- Save to: `<workspace_root>/src/classes/`
- Include interface if making class immutable
- Add corresponding test file in `<workspace_root>/src/tests/`
- Update tracking.md with new files created

**Step 3.4 - Testing:**
- Always add tests for new classes
- Use TestHelpers.bas to reduce redundancy
- Include failure management in all tests

---

### Phase 4: Post-Processing (MANDATORY)

**Step 4.1 - Line Endings:**
```bash
unix2dos <modified_file.cls>
unix2dos <modified_file.bas>
```
Run this for EVERY file you modify or create.

**Step 4.2 - Update Tracking:**
Update `.obt/tracking.md`:
- Add `[DONE]` to completed task bullets
- Update State section with what was accomplished
- Document what's remaining if task is incomplete

**Step 4.3 - Document Progress:**
If you stop mid-implementation:
- Update tracking.md State section with EXACTLY where you stopped
- List next steps clearly so work can be resumed
- Note any blockers or decisions needed

---

## Key Constraints (Never Violate These)

### 🚫 Prohibited Actions:
1. **Never modify:** DictionaryTestFixture.bas
2. **Never use:** External DLLs or ScriptingDictionary (macOS incompatible)
3. **Never return:** Partial files, diffs, or "only changed sections"
4. **Never skip:** unix2dos post-processing
5. **Never ignore:** implementations.md specifications

### ✅ Required Actions:
1. **Always read:** project-rules.md at start of EVERY VBA task
2. **Always return:** Complete, full file contents
3. **Always update:** tracking.md with progress
4. **Always add:** Tests for new classes
5. **Always use:** BetterArray for collections (not Dictionary)

---

## Error Handling Philosophy

From project-rules.md Section 7:
- Prefer explicit checks over `On Error Resume Next`
- If `On Error` is used, it MUST be:
  - Scoped to specific operations
  - Commented with reason
  - Restored with `On Error GoTo 0`
- Use ProjectError pattern for error management
- Silent failure is NEVER acceptable

---

## Output Format

Your output must be:
- ✅ Deterministic and reproducible
- ✅ Ready to paste directly into VBA editor
- ✅ Free of placeholders or TODOs (unless requested)
- ✅ Valid VBA syntax with no markdown inside code
- ❌ No emojis in code
- ❌ No explanations inside files (only as comments)

---

## When in Doubt

If an instruction is ambiguous:
1. Choose the MOST CONSERVATIVE interpretation
2. Preserve existing behavior
3. Ask the user for clarification before implementing
4. Refer back to project-rules.md for guidance

**Correctness beats cleverness.**

---

## Quick Reference Commands

**Starting a new task:**
```
1. Read: obt-skill/project-rules.md
2. Read: .obt/implementations.md
3. Read: .obt/tracking.md
4. Plan implementation
5. Execute changes
6. Run unix2dos
7. Update tracking.md
```

**File locations:**
- Source classes: `src/classes/*.cls`
- Test files: `src/tests/*TestFixture.bas`
- Task specs: `.obt/implementations.md`
- Progress tracking: `.obt/tracking.md`
- Coding standards: `obt-skill/project-rules.md`

---

## Skill Invocation

This skill is automatically invoked when VBA-related keywords are detected. You can also explicitly invoke it:

```
/obt <task description>
```

Example:
```
/obt refactor the ReportBuilder class to improve error handling
```

---

## Version History

- **1.0** (2026-02-04): Initial skill creation with project-rules architecture and .obt workspace

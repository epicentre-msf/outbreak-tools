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
Read: <workspace_root>/.obt/context-history.md
```
- `implementations.md` = Current goal, behaviors, and rules (user-set specification)
- `tracking.md` = Detailed task breakdown and progress tracking (tactical)
- `context-history.md` = Accumulated knowledge and learnings (strategic memory)

**Why context-history.md matters:**
This file prevents context loss over many iterations by preserving:
- Important decisions and rationale
- Patterns that work (and patterns that failed)
- Architectural choices
- Key learnings and gotchas

**Step 1.3 - Understand Workspace Structure:**
```
<workspace_root>/
├── src/
│   ├── classes/             ← VBA .cls files (organized in subfolders by topic)
│   │   ├── topic1/          ← Classes for specific topic
│   │   ├── topic2/          ← Classes for specific topic
│   │   └── stale/           ← Classes needing work/deletion (CHECK BEFORE USE)
│   ├── modules/             ← VBA .bas orchestration modules (same subfolder structure)
│   │   ├── topic1/          ← Modules for specific topic
│   │   ├── topic2/          ← Modules for specific topic
│   │   └── stale/           ← Modules needing work/deletion (CHECK BEFORE USE)
│   ├── tests/               ← Test files
│   └── [legacy folders]     ← DO NOT TOUCH - User will handle these
├── .obt/
│   ├── implementations.md   ← Current task specification (user-set)
│   ├── tracking.md          ← Current progress tracking (tactical)
│   └── context-history.md   ← Accumulated knowledge (strategic memory)
└── obt-skill/
    ├── SKILL.md             ← This file
    └── project-rules.md     ← Coding standards (READ FIRST)
```

---

### Phase 2: Planning (Before Any Code Changes)

**Step 2.1 - Analyze Request:**
- What files need to be modified?
- What new files need to be created?
- Does this align with `implementations.md`?
- What's the current state in `tracking.md`?

**Step 2.2 - Check for Stale Classes/Modules (CRITICAL):**
⚠️ **BEFORE working on ANY class or module, check its location:**

If the file is in a `stale/` subfolder:
1. **STOP immediately**
2. **ASK the user:** "The file [filename] is in the stale folder. Is this class/module ready to use, or is it in stale mode (needs deletion/refactoring)?"
3. **WAIT for user response** before proceeding
4. Only proceed if user confirms it's ready to use

**Why:** Stale folders contain classes/modules that may need deletion, major refactoring, or are incomplete. Never assume they're production-ready.

**Step 2.3 - Check Constraints:**
- Review project-rules.md Section 2.4 for files that MUST NOT be modified
- Verify the change fits within existing architecture (Section 4 of project-rules.md)
- Ensure no circular dependencies will be created
- **NEVER modify files in legacy folders** (only user can handle these)

**Step 2.4 - Update Tracking Plan:**
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
- **REMEMBER:** A "class" typically means TWO files (interface + implementation)
- **Default behavior:** Create both `IClassName.cls` (interface) and `ClassName.cls` (implementation)
- **Ask user which topic subfolder** to use (e.g., `src/classes/reporting/`, `src/classes/validation/`)
- If user doesn't specify, ask: "Which subfolder should I create this class in? (e.g., reporting, validation, etc.)"
- Save to: `<workspace_root>/src/classes/<topic>/IClassName.cls` AND `<workspace_root>/src/classes/<topic>/ClassName.cls`
- **Exception:** Only skip interface creation if it adds no value (simple data holder, internal helper)
- **If skipping interface:** Inform user with brief justification (e.g., "Creating DataHolder.cls without interface - simple data container")
- Add corresponding test file in `<workspace_root>/src/tests/` (test the implementation class)
- Update tracking.md with new files created

**Examples:**
- Request: "Create Validator class" → Creates `IValidator.cls` + `Validator.cls` (+ test)
- Request: "Split Processor into Parser and Writer" → Creates 4 files: `IParser.cls`, `Parser.cls`, `IWriter.cls`, `Writer.cls` (+ tests)

When creating NEW modules:
- **Ask user which topic subfolder** to use (same structure as classes)
- Save to: `<workspace_root>/src/modules/<topic>/ModuleName.bas`
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
- Update State section with timestamp (YYYY-MM-DD HH:MM format)
- Document what was accomplished and what remains
- Include hour and minute in all timestamps

**Step 4.3 - Update Context History (CRITICAL):**

Update `.obt/context-history.md` when ANY of these triggers occur:

**MANDATORY TRIGGERS:**
1. ✅ **After completing a significant task** (marked [DONE] in tracking.md)
2. ✅ **When making important architectural decisions**
3. ✅ **When solving significant problems or overcoming challenges**
4. ✅ **When discovering patterns that work well**
5. ✅ **When discovering gotchas or anti-patterns**
6. ✅ **At end of work session** (when stopping work)

**How to add an entry:**
```markdown
### [YYYY-MM-DD HH:MM] Brief Session Title (3-5 words)
**What:** One sentence describing what was done
**Why:** One sentence explaining the motivation
**Key Decision:** Most important choice made (if applicable)
**Challenge:** Main obstacle encountered (if applicable)
**Outcome:** ✅ success / ⚠️ partial / ❌ blocked
```

**What to include:**
- ✅ Decisions and rationale
- ✅ Challenges overcome
- ✅ Patterns discovered
- ✅ Gotchas found
- ✅ Results/outcomes
- ❌ NOT detailed code (that's in files)
- ❌ NOT step-by-step details (that's in tracking.md)

**Keep entries concise:** 3-5 sentences maximum per session entry.

**FILE SIZE MONITORING (CRITICAL):**
After adding a new session entry, check file health:
- Count Recent Sessions entries
- ⚠️ **If 25-30 entries:** Archive older sessions (keep most recent 10)
- 🚨 **If 35+ entries:** Immediate archiving required
- Check if file is approaching ~1500 lines (use `wc -l`)
- If WARNING threshold reached, follow archiving strategy in project-rules.md Section 2.4

**Also update Knowledge Base section when:**
- You establish a new architectural pattern
- You discover a failed approach worth documenting
- You find an important gotcha
- You learn a cross-platform consideration

**Step 4.4 - Document Progress:**
If you stop mid-implementation:
- Update tracking.md State section with EXACTLY where you stopped
- List next steps clearly so work can be resumed
- Note any blockers or decisions needed
- Add entry to context-history.md summarizing the session

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
- Source classes: `src/classes/<topic>/*.cls` (organized by topic in subfolders)
- Source modules: `src/modules/<topic>/*.bas` (organized by topic in subfolders)
- Stale classes: `src/classes/stale/*.cls` (CHECK before using)
- Stale modules: `src/modules/stale/*.bas` (CHECK before using)
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

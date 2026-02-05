# OutbreakTools Project Rules & Coding Standards

**Version:** 1.0
**Last Updated:** 2026-02-04
**Authority:** This document is the authoritative source for all OutbreakTools VBA development standards

---

## Section 1: Scope and Role

### 1.1 Your Role as VBA Assistant

You are a **VBA code assistant** focused on production-ready code.

**Permitted Actions:**
- ✅ Analyze existing VBA code
- ✅ Improve correctness, robustness, and clarity
- ✅ Refactor large portions of code
- ✅ Create new classes and modules
- ✅ Produce error-proof, bug-free code
- ✅ Add comprehensive code comments for easy handover
- ✅ Respect architectural and stylistic constraints

**Prohibited Actions:**
- ❌ Invent new features unless explicitly requested
- ❌ Simplify logic in ways that change behavior
- ❌ Remove existing or legacy code unless explicitly requested
- ❌ Move code between repositories unless explicitly requested
- ❌ Act as a tutor or provide explanations instead of code

---

## Section 2: File Handling Rules (CRITICAL)

### 2.1 Full File Rewrites Only ⚠️

**MANDATORY RULE:**
When modifying ANY code file, you MUST:
- Return the FULL rewritten file (never diffs, patches, or snippets)
- NEVER say "only the changed part is shown"
- NEVER assume the user will merge changes manually
- NEVER return unfinished improvements
- NEVER use placeholders like `// ... rest of code`

**Even if a single line changes, return the entire file.**

### 2.2 Workflow Tracking

**Tracking File Location:** `<workspace_root>/.obt/tracking.md`

**Required Tracking Structure:**
```markdown
Goal
- [High-level objective of the current task]

Tasks
- [Bullet point for sub-task 1]
- [DONE] [Completed sub-task with DONE marker]
- [Bullet point for sub-task 3]

State
- Latest update (YYYY-MM-DD): [What was accomplished, what remains]
```

**Tracking Rules:**
- For each big task, draft planned implementations as bullets in tracking.md
- ALWAYS track progress by adding `[DONE]` to completed task bullets
- NEVER leave tasks in unfinished/stale state without documenting exactly what's next
- ALWAYS read tracking.md entirely before implementing new changes
- If you stop mid-implementation, provide detailed information in State section

### 2.3 Implementation Specifications

**Implementation File Location:** `<workspace_root>/.obt/implementations.md`

This file contains:
- Goal: What is the main task
- Behaviors: Expected behaviors and specifications
- Rules: Specific rules for this implementation

**Rules:**
- Re-read implementations.md if necessary to know scope of work
- NEVER modify implementations.md (it's set by the user)

### 2.4 File Modification Scope

**Rules:**
- Modify ONLY the file explicitly requested
- Do not touch other files unless explicitly instructed
- Do not describe hypothetical changes to other files

### 2.5 Specific File Exceptions ⚠️

**CRITICAL - NEVER MODIFY:**
- `DictionaryTestFixture.bas`

---

## Section 3: VBA Language Constraints

### 3.1 Platform & Version

**Target Platform:**
- Excel VBA for Windows AND macOS
- No external DLL libraries (macOS incompatibility)
- No external references unless explicitly mentioned
- Late binding preferred over early binding unless stated otherwise

### 3.2 Explicitness Requirements

**MANDATORY in all modules:**
```vba
Option Explicit
```

**Variable Declaration:**
- All variables MUST be explicitly declared
- Avoid implicit `Variant` usage unless intentional and documented
- Declare types explicitly for clarity

---

## Section 4: Project Architecture

### 4.1 Project Structure

```
src/
├── classes/                    ← Business logic (.cls files)
│   ├── <topic1>/              ← Classes organized by topic/feature
│   ├── <topic2>/              ← Classes organized by topic/feature
│   ├── stale/                 ← Classes needing work/deletion (CHECK BEFORE USE)
│   └── ...
│
├── modules/                    ← Orchestration modules (.bas files)
│   ├── <topic1>/              ← Modules organized by topic/feature (mirrors classes structure)
│   ├── <topic2>/              ← Modules organized by topic/feature
│   ├── stale/                 ← Modules needing work/deletion (CHECK BEFORE USE)
│   └── ...
│
├── tests/                      ← Test fixtures and test helpers
│   └── *TestFixture.bas       ← Test files
│
└── [legacy folders]/           ← DO NOT TOUCH - User manages these
```

**File Types:**
- `.cls` files: Business logic classes (in `src/classes/<topic>/`)
- `.bas` files: Orchestration modules (in `src/modules/<topic>/`) and test helpers
- Interfaces: Named with `I` prefix (e.g., `IFoo`)

**Folder Organization:**
- Classes and modules are organized in subfolders by topic/feature area
- Each topic subfolder contains related classes or modules
- `stale/` subfolders contain files that may need deletion, major refactoring, or are incomplete
- Legacy folders in `src/` are managed by user only - NEVER modify these

### 4.2 Stale Mode Checking (CRITICAL)

⚠️ **BEFORE working on ANY file in a `stale/` subfolder:**

1. **STOP and ASK the user:** "The file [filename] is in the stale folder. Is this class/module ready to use, or is it in stale mode (needs deletion/refactoring)?"
2. **WAIT for explicit user confirmation** before proceeding
3. Only proceed if user confirms the file is ready to use

**Why:** Stale folders contain classes/modules that may:
- Need to be deleted entirely
- Require major refactoring
- Be incomplete or non-functional
- Be outdated and no longer used

**Never assume stale files are production-ready.**

### 4.3 Architectural Principles

**You MUST:**
- ✅ Respect existing class responsibilities
- ✅ Avoid circular dependencies
- ✅ Avoid moving logic across layers without instruction
- ✅ Ensure new classes fit in overall project structure (`src/classes/<topic>/`)
- ✅ **Ask user which topic subfolder** to use when creating new classes/modules
- ✅ Add new classes in `src/classes/<topic>/`
- ✅ Add new modules in `src/modules/<topic>/`
- ✅ Add new tests in `src/tests/`
- ✅ Always add tests for newly created classes
- ✅ Use TestHelpers.bas in tests to reduce redundancy
- ✅ Use existing classes if required - don't reinvent the wheel
- ✅ Aim for efficiency - code should execute fast

**You MUST NOT:**
- ❌ **NEVER modify files in legacy folders within `src/`** (only user manages these)
- ❌ Never delete legacy folders or their contents
- ❌ Never move files into or out of legacy folders
- ❌ Never assume files in `stale/` subfolders are ready for use

**Class Design:**
- Most classes have a dedicated interface for immutability
- Always keep interface implementation at the END of class code

**Module Design:**
- Modules orchestrate classes and provide helper functions
- Modules mirror the same topic structure as classes

---

## Section 5: Naming Conventions (CRITICAL)

### 5.1 Variables and Members

**Casing:**
- `camelCase` for variables and parameters
  ```vba
  Dim rowCount As Long
  Dim userName As String
  ```
- `PascalCase` for classes and public members
  ```vba
  Public Function GetRowCount() As Long
  Public Property Get TableName() As String
  ```

**Restrictions:**
- ❌ No Hungarian notation (`strName`, `lngCount`)
- ❌ No single-letter variables except tight loops (`i`, `j`, `k`)
- ❌ Avoid ambiguous names: `sheet`, `workbook` (clashes with VBA object model)
- ❌ CRITICAL: Never name variables identically to functions/subs in the same file
  ```vba
  ' ❌ BAD - VBA will freeze/crash
  Dim devSheet As Worksheet
  Set devSheet = DevSheet()  ' Function name collision!

  ' ✅ GOOD
  Dim targetSheet As Worksheet
  Set targetSheet = DevSheet()
  ```

**String Handling:**
- Use `Chr(34)` for double quotes in strings
- Or escape correctly using VBA syntax (`""` for embedded quotes)
- Avoid syntax errors from improper quote handling

### 5.2 Procedure Naming

**Pattern:**
- Verbs for actions: `BuildTable`, `ParseFile`, `CalculateTotal`
- Nouns for getters: `RowCount`, `ProjectRoot`, `ColumnIndex`

**Avoid:**
- ❌ Vague names: `HandleStuff`, `DoThing`, `ProcessData`

### 5.3 Platform Constraints

**CRITICAL - NEVER USE:**
- ❌ External DLLs (macOS incompatible)
- ❌ ScriptingDictionary (macOS incompatible)

**Instead, use:**
- ✅ BetterArray for collections and dictionary-like functionality

---

## Section 6: Documentation Standards

### 6.1 Documentation Tags

This project uses **structured comment tags** (similar to roxygen).

**Common Tags:**
```vba
' @description Brief description of what this does
' @details Longer explanation if needed
' @param paramName Description of parameter
' @return Description of return value
' @export Marks if this should be exported/public
' @label Classification or category label
```

**Rules:**
- ✅ Preserve existing documentation tags
- ✅ Add documentation ONLY as comments
- ✅ Clean, document, and comment class interfaces
- ✅ Add section headers to lengthy classes
- ❌ Never remove or rename tags
- ❌ Do not invent new tags unless instructed

### 6.2 Code Comments

**Purpose:**
- Enable easy handover to other developers
- Explain WHY, not just WHAT
- Document non-obvious logic

**Example:**
```vba
' Check if worksheet exists before export to avoid data loss
If WorksheetExists(targetBook, sheetName) Then
    ' Import from existing sheet instead of overwriting
    ImportFrom targetBook
    Exit Sub
End If
```

---

## Section 7: Error Handling Philosophy

### 7.1 Error Handling Pattern

**Preferred Approach:**
- Use explicit checks over `On Error Resume Next`
- Implement error handling with ProjectError pattern

**If Using `On Error`:**
```vba
' Must be scoped to specific operation
On Error Resume Next
Set ws = wb.Worksheets(sheetName)
On Error GoTo 0  ' ALWAYS restore error handling

If ws Is Nothing Then
    ' Handle the error explicitly from ProjectError in IChecking/Checking class
    Err.Raise ObjectNotItialized,  <classname>, <message>
End If
```

There is a short ThrowError Sub in most of the classes:

```vba
'@sub-title Raise a ProjectError-based exception.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
```

**Requirements:**
- ✅ Error handling must be scoped
- ✅ Must be commented with reason
- ✅ Must be restored with `On Error GoTo 0`
- ❌ Silent failure is NEVER acceptable

### 7.2 Validation and Checks

- Add checks and validation where appropriate
- Provide notification for check failures
- Use ProjectError for consistent error management

---

## Section 8: Style and Formatting

### 8.1 Code Formatting

**Indentation:**
- 4 spaces (NOT tabs)
- Consistent indentation depth

**Structure:**
- One statement per line
- No chained logic for brevity's sake
- Readability > cleverness

**Example:**
```vba
' ✅ GOOD - Clear and readable
Dim result As Long
result = CalculateValue(param1)
result = result * 2
If result > 100 Then
    ProcessResult result
End If

' ❌ BAD - Chained for brevity
If CalculateValue(param1) * 2 > 100 Then ProcessResult CalculateValue(param1) * 2
```

### 8.2 Formatting Preservation

**Rule:**
Do NOT reformat code unless logic is being modified.

Preserve existing formatting style in files you're not changing.

---

## Section 9: Testing Requirements

### 9.1 Test Standards

**Requirements:**
- ✅ Always add tests for newly created classes
- ✅ Use TestHelpers.bas to reduce redundancy
- ✅ ALWAYS add failure management in tests
- ✅ Test files go in `src/tests/`
- ✅ Follow naming convention: `*TestFixture.bas`
- ✅ Test files should have same architecture as in other tests, based on CustomTest. Read `TestLLChoices` for inspirations and learning on overall test implementations

**Test Structure Example:**
```vba


Public Sub TestMethodNameScenarioExpectedResult()
    CustomTestSetTitles Assert, "<classname>", "<Testmethodname>"
    On Error GoTo TestFail

    ' Arrange
    Dim sut As ClassName
    Set sut = class.Create() 'You can also leverage module Initialise for repeating objects


    ' Act
    Dim result As Variant
    result = sut.MethodToTest(param)

    ' Assert
    Assert.AreEqual result, expectedValue, "<message>"

    Exit Sub
TestFail:
    Err.Raise Err.Number, "TestMethodName", Err.Description
End Sub
```

---

## Section 10: Assumptions You Must NOT Make

**Never Assume:**
- ❌ The project can use DLLs or external libraries
- ❌ Windows-only APIs are available
- ❌ Classes can depend on module code (modules can use classes, not vice versa)

**You CAN Assume:**
- ✅ Classes can use other classes
- ✅ Modules can use classes
- ✅ BetterArray is available for collection needs

**Principle:**
Preserve intent first, elegance second.

---

## Section 11: Output Expectations

### 11.1 Code Quality

Your output must be:
- ✅ Deterministic - same input produces same output
- ✅ Reproducible - can be run repeatedly
- ✅ Ready to paste directly into VBA editor
- ✅ Free of placeholders or TODOs (unless explicitly requested)
- ✅ Syntactically valid VBA

### 11.2 Output Restrictions

**Never include:**
- ❌ Emojis in code or comments
- ❌ Markdown formatting inside code files
- ❌ Explanations inside the file (only comments)
- ❌ Placeholder comments like `' TODO: implement this`
- ❌ Incomplete implementations

---

## Section 12: Post-Processing (MANDATORY)

### 12.1 Line Ending Conversion

**ALWAYS run unix2dos after modifying files:**
```bash
unix2dos <filename.cls>
unix2dos <filename.bas>
```

This ensures proper line endings for VBA editor compatibility.

### 12.2 Final Checklist

Before delivering code:
- [ ] Full file returned (not diff/snippet)
- [ ] unix2dos executed
- [ ] tracking.md updated with progress
- [ ] Tests added for new classes
- [ ] All naming conventions followed
- [ ] Error handling implemented
- [ ] Documentation tags present
- [ ] No syntax errors

---

## Section 13: Specific Coding Rules

### 13.1 Type Checking

**Use TypeName instead of TypeOf:**
```vba
' ✅ GOOD
If TypeName(obj) = "Worksheet" Then

' ❌ AVOID
If TypeOf obj Is Worksheet Then
```

### 13.2 Collections

**Use BetterArray:**
```vba
' ✅ GOOD - Cross-platform
Dim items As BetterArray
Set items = New BetterArray

' ❌ BAD - Not available on macOS
Dim items As Dictionary
Set items = New Dictionary
```

### 13.3 ListObject Creation

**Correct Syntax:**
```vba
' ✅ GOOD
Set lo = ws.ListObjects.Add( _
    SourceType:=xlSrcRange, _
    Source:=dataRange, _
    XlListObjectHasHeaders:=xlYes _
)

' ❌ BAD - xlSrcRange is not a parameter name
Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
```

**Parameters:**
1. `SourceType` (Long)
2. `Source` (Range or String)
3. `LinkSource` (Optional Boolean)
4. `XlListObjectHasHeaders` (Optional constant)
5. `Destination` (Optional Range)

---

## Section 14: Decision-Making Guidelines

### 14.1 When in Doubt

If an instruction is ambiguous:
1. Choose the MOST CONSERVATIVE interpretation
2. Preserve existing behavior
3. Avoid creative refactors
4. Ask the user for clarification before implementing

### 14.2 Priority Order

When rules conflict:
1. **Correctness** (code must work)
2. **Safety** (don't break existing functionality)
3. **Architecture** (respect project structure)
4. **Style** (follow conventions)
5. **Elegance** (nice to have, lowest priority)

**Principle:**
Correctness beats cleverness. Always.

---

## Section 15: Quick Reference

### File Locations
- **Source classes:** `src/classes/<topic>/*.cls` (organized in subfolders by topic)
- **Source modules:** `src/modules/<topic>/*.bas` (organized in subfolders by topic)
- **Stale classes:** `src/classes/stale/*.cls` (⚠️ CHECK with user before using)
- **Stale modules:** `src/modules/stale/*.bas` (⚠️ CHECK with user before using)
- **Tests:** `src/tests/*TestFixture.bas`; `src/tests/*/TestLLChoices.bas`
- **Test helpers:** `src/tests/TestHelpers.bas`
- **Legacy folders:** `src/[legacy folders]/` (❌ DO NOT MODIFY - user manages)
- **Tracking:** `.obt/tracking.md`
- **Specifications:** `.obt/implementations.md`

### Common Patterns
```vba
' Option Explicit (mandatory)
Option Explicit

' Error handling with ProjectError
On Error GoTo ErrorHandler
' ... code ...
Exit Sub
ErrorHandler:
    Err.Raise ProjectError.<ErrorContext>, "MethodName", "Error message"
End Sub

' String with quotes
Dim sql As String
sql = "SELECT * FROM table WHERE name = " & Chr(34) & "value" & Chr(34)
```

### Pre-flight Checklist
1. Read `project-rules.md` (this file)
2. Read `.obt/implementations.md`
3. Read `.obt/tracking.md`
4. Plan changes
5. Implement with full file return
6. Run unix2dos
7. Update tracking.md
8. Add tests if needed

---

## Version History

- **1.0** (2026-02-04): Initial project rules extracted from legacy instructions.md

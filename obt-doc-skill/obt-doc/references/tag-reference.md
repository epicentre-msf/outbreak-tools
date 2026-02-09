# OBT Documentation Tag Reference

This file is the authoritative grammar specification for all documentation tags
used in OutbreakTools VBA class files. An automated HTML documentation generator
will rely on these exact patterns, so consistency is critical.

---

## Table of Contents

1. [General Syntax Rules](#1-general-syntax-rules)
2. [Class-Level Tags](#2-class-level-tags)
3. [Rubberduck Annotation Tags](#3-rubberduck-annotation-tags)
4. [Section Tags](#4-section-tags)
5. [Member-Level Tags](#5-member-level-tags)
6. [Parameter and Return Tags](#6-parameter-and-return-tags)
7. [Error and Dependency Tags](#7-error-and-dependency-tags)
8. [Supplementary Tags](#8-supplementary-tags)
9. [Deprecated / Non-Canonical Tags](#9-deprecated--non-canonical-tags)
10. [Parser Extraction Guide](#10-parser-extraction-guide)

---

## 1. General Syntax Rules

All documentation tags are VBA comment lines that start with `'@`. The parser
identifies a doc block as a contiguous group of comment lines that precede a
`Public`, `Private`, or `Friend` declaration.

**Line format:**
```
'@tagname <value>
```

**Multi-line content:**
Tags whose value spans multiple lines use continuation comment lines. The
continuation line starts with `'` (a plain comment) and is indented to align
with the text above. The block ends when the next `'@tag` or a blank line or
a code line is encountered.

```vba
'@details
'This method scans the dictionary for matching entries. When no match
'is found, it returns an empty BetterArray rather than Nothing, which
'lets callers iterate safely without a Nothing-check.
'@param colName String. Column header to search for.
```

**Tag ordering within a doc block:**
Tags should follow this order (omit tags that do not apply):

1. `@label:<id>` or `@jump:<id>`
2. `@sub-title` or `@prop-title`
3. `@details`
4. `@param` (one per parameter, in signature order)
5. `@return`
6. `@throws` (one per error condition)
7. `@depends`
8. `@export`
9. `@remarks`
10. `@note`
11. `@todo`

---

## 2. Class-Level Tags

These tags appear once per file, in the header block after `Option Explicit`
and after the Rubberduck annotations.

### @class

**Syntax:** `'@class <ClassName>`
**Required:** Yes (implementation and interface files)
**Purpose:** Names the class being documented. This is the anchor the HTML
generator uses to build the page title and URL slug.

```vba
'@class LLdictionary
```

For interface files, use the interface name:
```vba
'@class ILLdictionary
```

### @description (class level)

**Syntax:**
```vba
'@description
'<Multi-line overview paragraph.>
```
**Required:** Yes
**Purpose:** A 2-5 sentence explanation of what the class does, who consumes
it, and how it fits into the architecture. Written for a beginner.

```vba
'@description
'The LLdictionary class centralises all logic that manipulates the Excel
'"Dictionary" worksheet used across setup and designer workflows. It wraps
'a DataSheet for consistent range access and provides methods for column
'management, row operations, preparation, import/export, and translation.
'Consumers interact with a single surface area via the ILLdictionary
'interface, and tests target the unified object model.
```

### @depends (class level)

**Syntax:** `'@depends <Class1>, <Class2>, <Class3>`
**Required:** When the class instantiates or consumes other project classes.
**Purpose:** Lists all project-level dependencies. Helps newcomers understand
the dependency graph and helps the HTML generator build cross-links.

```vba
'@depends DataSheet, Checking, HiddenNames, BetterArray
```

### @version

**Syntax:** `'@version <version-string or date>`
**Required:** Optional
**Purpose:** Tracks the documented version.

```vba
'@version 1.0 (2026-02-09)
```

### @author

**Syntax:** `'@author <name>`
**Required:** Optional
**Purpose:** Credits the original author.

```vba
'@author Yves Amevoin
```

---

## 3. Rubberduck Annotation Tags

These are consumed by the Rubberduck VBA add-in for code navigation. They are
NOT documentation tags per se, but must be preserved and placed correctly
because the codebase relies on them.

### @PredeclaredId

**Syntax:** `'@PredeclaredId`
**Where:** After `Option Explicit`, before class-level doc tags.
**Purpose:** Marks the class as having a default (predeclared) instance, which
enables the `ClassName.Create(...)` factory pattern.

### @Interface

**Syntax:** `'@Interface`
**Where:** Interface files only, after `Option Explicit`.
**Purpose:** Marks the file as an interface contract.

### @Folder

**Syntax:** `'@Folder("<FolderName>")`
**Where:** After `Option Explicit`.
**Purpose:** Organises the class in Rubberduck's Code Explorer.

```vba
'@Folder("Dictionary")
```

### @ModuleDescription

**Syntax:** `'@ModuleDescription("<one-line summary>")`
**Where:** After `@Folder`.
**Purpose:** Provides a short description visible in Rubberduck's UI. This
should match the first sentence of `@description`.

```vba
'@ModuleDescription("Linelist dictionary utilities")
```

### @IgnoreModule

**Syntax:** `'@IgnoreModule <Inspection1>, <Inspection2>, ...`
**Where:** After `@ModuleDescription`.
**Purpose:** Suppresses Rubberduck code inspections at the module level.

```vba
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing
```

### @Ignore

**Syntax:** `'@Ignore <InspectionName>`
**Where:** Immediately above the line triggering the inspection.
**Purpose:** Suppresses a single Rubberduck inspection.

---

## 4. Section Tags

### @section

**Syntax:** `'@section <SectionName>`
**Required:** Yes -- every file should have at least one section.
**Purpose:** Groups related members under a named heading. The HTML generator
uses sections as navigation anchors within a class page.

Convention: follow the `@section` line with a visual divider:

```vba
'@section Instantiation
'===============================================================================
```

Optionally add a `@description` for non-obvious sections:

```vba
'@section ExportCounterManagement
'===============================================================================
'@description Validation, caching, and persistence of the export column counter.
```

**Naming guidelines:**
- Use PascalCase or Title Case: `Instantiation`, `Data Access`, `Row Management`
- Keep names short (1-3 words)
- Common section names: `Instantiation`, `Data Elements`, `Column Operations`,
  `Row Operations`, `Preparation`, `Data Exchange`, `Specialised Views`,
  `Validation`, `Helpers`, `Interface Implementation`

---

## 5. Member-Level Tags

### @label:<id>

**Syntax:** `'@label:<unique-id>`
**Required:** Yes, on every documented member in implementation files.
**Purpose:** Assigns a stable, unique identifier to the member. Interface files
reference this via `@jump:<id>`. The HTML generator uses it as an anchor ID.

**Naming rules:**
- Use lowercase with hyphens: `create`, `add-rows`, `column-exists`
- Must be unique within the file.
- Should be short but descriptive.

```vba
'@label:create
'@label:column-exists
'@label:apply-exports
```

### @jump:<id>

**Syntax:** `'@jump:<id>`
**Required:** Yes, on every documented member in interface files.
**Purpose:** Cross-references the implementation's `@label`. The `<id>` must
match exactly.

```vba
'@jump:create
'@jump:column-exists
```

### @sub-title

**Syntax:** `'@sub-title <one-line imperative summary>`
**Required:** Yes, on `Sub` and `Function` members.
**Purpose:** A brief (one sentence, imperative verb) summary of what the
member does. Shown in the HTML index alongside the member name.

```vba
'@sub-title Create a dictionary object
'@sub-title Append a column at the end of the dictionary
'@sub-title Sort the dictionary on sheet name and main section
```

### @prop-title

**Syntax:** `'@prop-title <one-line summary>`
**Required:** Yes, on `Property Get/Let/Set` members.
**Purpose:** Same as `@sub-title` but for properties.

```vba
'@prop-title DataSheet backing store
'@prop-title Number of export columns
'@prop-title Dictionary preparation status
```

### @details

**Syntax:**
```vba
'@details
'<Multi-line explanation paragraph.>
```
**Required:** Yes on all `Public` members. Recommended on non-trivial
`Private` members.
**Purpose:** A longer explanation of what the member does, covering:
- Behaviour description (what it does step by step)
- Edge cases (what happens with empty inputs, Nothing, zero-length strings)
- Side effects (does it modify the worksheet? persist values? reset caches?)
- Why it exists (context for a newcomer)

Write for a beginner. Avoid jargon without explanation. If the method
delegates to another class, mention that relationship.

```vba
'@details
'Creates a linelist dictionary wrapper around the supplied worksheet. The
'routine initialises a backing DataSheet so all subsequent access goes through
'a consistent abstraction and callers interact with a predeclared instance.
'When no export count is supplied, the method falls back to the class
'constant DEFAULTNUMBEROFEXPORTS (20).
```

### @export

**Syntax:** `'@export`
**Required:** On every member that is intentionally called from outside the
class -- not just factory methods.
**Purpose:** Signals to the HTML generator and to human readers that this member
is part of the class's callable API surface. The parser uses `@export` to build
the "Public API" index on each class page, separating outward-facing members
from internal helpers that happen to be `Public` (e.g. `Self`, property setters
used only during `Create`).

**When to apply `@export`:**
- Factory methods (`Create`)
- Public operations invoked by modules or other classes (`AddRows`, `RemoveRows`,
  `Sort`, `Import`, `Export`, `Translate`, `Prepare`, `Clean`, etc.)
- Any `Public` member that external consumers are expected to call

**When NOT to apply `@export`:**
- `Self` property (internal wiring for the factory pattern)
- `Property Set/Let` members used only by `Create` during initialisation
- Members that are `Public` solely to satisfy VBA's `With New` pattern but are
  not part of the intended caller contract

**Examples from the codebase:**

```vba
'@label:create
'@sub-title Create an Analysis instance
'@export
Public Function Create(ByVal hostsheet As Worksheet) As IAnalysis

'@label:addrows
'@sub-title Add rows based on worksheet header selection
'@export
Private Sub AddRows()

'@label:sort
'@sub-title Sort analysis listobjects and enforce minimum rows
'@export
Private Sub Sort()

'@label:import
'@sub-title Import analysis tables from an external worksheet
'@export
Private Sub Import(ByVal sourceSheet As Worksheet)

'@label:export
'@sub-title Export the analysis worksheet to a workbook
'@export
Private Sub Export(ByVal sourceWb As Workbook)

'@label:translate
'@sub-title Translate analysis labels using the provided engine
'@export
Private Sub Translate(ByVal TransObject As ITranslationObject)
```

**Note:** In the OutbreakTools architecture many exported members are declared
`Private` in the implementation class because callers go through the interface
(`IClassName`). The `@export` tag marks intent -- "this is part of the public
contract" -- regardless of the VBA visibility keyword.

---

## 6. Parameter and Return Tags

### @param

**Syntax:** `'@param <paramName> <Type>. <Description.>`
**Required:** Yes, one tag per parameter on any member that accepts arguments.
**Purpose:** Documents a single parameter.

**Format rules (critical for parser):**
- `<paramName>` must exactly match the VBA signature name (case-sensitive).
- `<Type>` is the VBA data type. For optional parameters, prepend `Optional`:
  `Optional String`, `Optional Long`, `Optional Boolean`.
- End the description with a period.
- For parameters with default values, state the default:
  `Defaults to "__all__".`
- For `ByRef` parameters that get modified, state it:
  `ByRef Long. Populated with the row count on return.`
- Parameters must be listed in the same order as the VBA signature.

**Examples:**

```vba
'@param dictWksh Worksheet. The worksheet hosting the dictionary data.
'@param dictStartRow Long. Header row of the dictionary (1-based).
'@param dictStartColumn Long. First column of the dictionary (1-based).
'@param numberOfExports Optional Long. Number of export columns to configure. Defaults to 20.
```

```vba
'@param colName String. Column header to search for.
'@param checkValidity Optional Boolean. When True, validates the column against the allowed schema. Defaults to False.
```

```vba
'@param targetCell Range. The selection whose height dictates the number of rows to insert.
'@param insertShift Optional Boolean. When True, worksheet rows are inserted to protect stacked tables. Defaults to True.
```

### @return

**Syntax:** `'@return <Type>. <Description.>`
**Required:** Yes, on every `Function` and `Property Get` that returns a value.
**Purpose:** Documents what the member hands back to the caller.

**Format rules:**
- `<Type>` is the VBA return type.
- End the description with a period.
- If the member can return `Nothing`, say when.

**Examples:**

```vba
'@return ILLdictionary. A fully initialised dictionary instance ready for use.
'@return Boolean. True when the column header exists in the dictionary.
'@return BetterArray. A collection of distinct values, empty if the column is missing.
'@return Long. The 1-based column index, or 0 when not found.
'@return Worksheet. The worksheet backing this dictionary, or Nothing when the data store is uninitialised.
```

---

## 7. Error and Dependency Tags

### @throws

**Syntax:** `'@throws <ErrorType> When <condition>.`
**Required:** On any member that calls `ThrowError`, `Err.Raise`, or delegates
to code that raises errors.
**Purpose:** Tells the caller what exceptions to guard against.

**Examples:**

```vba
'@throws ProjectError.ObjectNotInitialized When the worksheet object is Nothing.
'@throws ProjectError.InvalidArgument When the column name is empty or not found.
'@throws ProjectError.OutOfRange When the requested row exceeds the data boundary.
```

### @depends (member level)

**Syntax:** `'@depends <ClassA>, <ClassB>`
**Required:** Optional. Use when a specific method instantiates classes that
are not already listed in the class-level `@depends`.
**Purpose:** Helps trace instantiation paths.

```vba
'@depends CustomTable, BetterArray
```

---

## 8. Supplementary Tags

### @remarks

**Syntax:**
```vba
'@remarks <Text or multi-line block.>
```
**Required:** Optional.
**Purpose:** Implementation-level observations aimed at maintainers. Not part
of the API contract -- callers do not need this information.

```vba
'@remarks The sort relies on Excel's built-in Range.Sort, which is
'locale-sensitive. Always test on both Windows and macOS.
```

### @note

**Syntax:** `'@note <Text.>`
**Required:** Optional.
**Purpose:** Important caveats that callers should be aware of. Unlike
`@remarks`, these affect how you USE the member.

```vba
'@note This method mutates the worksheet in place. There is no undo.
'@note Column lookups are case-insensitive by default.
```

### @todo

**Syntax:** `'@todo <Description of outstanding work.>`
**Required:** Optional.
**Purpose:** Marks items that need future attention. The HTML generator can
compile a project-wide TODO list from these.

```vba
'@todo Clarify behaviour when the target sheet contains merged cells.
'@todo Add support for multi-column sort keys.
```

---

## 9. Deprecated / Non-Canonical Tags

The codebase contains some legacy tag variants. When documenting, convert
these to their canonical form.

| Legacy Tag | Canonical Replacement | Notes |
|------------|----------------------|-------|
| `@params` (block) | Individual `@param` lines | Split each parameter onto its own line |
| `@returned` | `@return` | Same syntax |
| `@returns` | `@return` | Same syntax |
| `@pram` | `@param` | Typo correction |
| `@fun-title` | `@sub-title` | Use `@sub-title` for both Sub and Function |
| `@hprefix` | `@param` | Convert to a standard `@param` tag |
| `@colName` | `@param colName` | Convert to standard form |
| `@includeIds` | `@param includeIds` | Convert to standard form |
| `@property` | `@prop-title` | Use the title variant |
| `@method` | `@sub-title` | Use the title variant |
| `@remark` (singular) | `@remarks` | Standardise to plural |

**Block-form @params conversion example:**

Before (legacy):
```vba
'@params
'- dictWksh Worksheet hosting the dictionary data
'- dictStartRow Long. Header row of the dictionary
```

After (canonical):
```vba
'@param dictWksh Worksheet. The worksheet hosting the dictionary data.
'@param dictStartRow Long. Header row of the dictionary (1-based).
```

---

## 10. Parser Extraction Guide

This section describes how a future automated tool should extract tags from
`.cls` files. The patterns are designed to be parseable with simple regex or
line-by-line scanning.

### Extraction Algorithm

```
FOR each .cls file:
  1. Extract class-level tags (between Option Explicit and first member)
     - @class, @description, @depends, @version, @author
     - @PredeclaredId, @Interface, @Folder, @ModuleDescription

  2. Scan for @section headers
     - Each @section starts a new group
     - Optional @description below the divider line

  3. For each member (Sub, Function, Property):
     a. Collect the doc block: all contiguous comment lines above the
        declaration that contain @tags or are continuation lines
     b. Extract: @label/@jump, @sub-title/@prop-title, @details,
        @param (multiple), @return, @throws (multiple), @depends,
        @export, @remarks, @note, @todo
     c. Associate the member with its enclosing @section

  4. Cross-reference @label and @jump between implementation and
     interface files to build linked documentation
```

### Regex Patterns

Tag line identification:
```
^'@(\w[\w-]*):?(.*)$
```

Param extraction:
```
^'@param\s+(\w+)\s+(.+)$
  Group 1: parameter name
  Group 2: type + description
```

Return extraction:
```
^'@return\s+(.+)$
  Group 1: type + description
```

Section extraction:
```
^'@section\s+(.+)$
  Group 1: section name
```

Label/Jump extraction:
```
^'@(label|jump):(\S+)$
  Group 1: "label" or "jump"
  Group 2: the identifier
```

### Output Structure (per class)

The parser should produce a structure like this (JSON shown for illustration):

```json
{
  "class": "LLdictionary",
  "file": "src/classes/dictionary/LLdictionary.cls",
  "interface": "src/classes/dictionary/ILLdictionary.cls",
  "description": "The LLdictionary class centralises all logic...",
  "depends": ["DataSheet", "Checking", "HiddenNames", "BetterArray"],
  "predeclaredId": true,
  "sections": [
    {
      "name": "Instantiation",
      "members": [
        {
          "label": "create",
          "name": "Create",
          "kind": "Function",
          "visibility": "Public",
          "title": "Create a dictionary object",
          "details": "Creates a linelist dictionary wrapper...",
          "params": [
            {"name": "dictWksh", "type": "Worksheet", "description": "..."},
            {"name": "dictStartRow", "type": "Long", "description": "..."}
          ],
          "return": {"type": "ILLdictionary", "description": "..."},
          "throws": [],
          "depends": ["DataSheet", "Checking"],
          "export": true
        }
      ]
    }
  ]
}
```

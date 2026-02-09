---
name: obt-doc
description: >
  Generate comprehensive, roxygen-style documentation for OutbreakTools VBA
  classes and interfaces. Use this skill whenever the user asks to document a
  .cls file, generate class documentation, add doc headers, create API docs for
  VBA code, or mentions "document", "documentation", "doc comments", "roxygen",
  or "HTML docs" in the context of VBA classes. Also trigger when the user says
  things like "make this class understandable", "add docs", "write headers", or
  references wanting to produce automated documentation from VBA source files.
  Always trigger on requests involving @param, @return, @section, @details tags
  or any structured comment-based documentation for .cls files.
---

# OutbreakTools VBA Class Documentation Skill (obt-doc)

**Version:** 1.0
**Purpose:** Produce complete, structured, machine-parseable documentation for
OutbreakTools VBA classes so that (a) a beginner VBA developer can understand
every class just by reading the doc comments, and (b) an automated tool can
later extract the tags into HTML reference pages (like roxygen for R or JSDoc
for JavaScript).

---

## Why This Skill Exists

The OutbreakTools codebase has 200+ VBA class files organised as
interface/implementation pairs (`IFoo.cls` + `Foo.cls`). The project already
uses documentation tags (`@section`, `@param`, `@details`, etc.) but coverage
is uneven: some methods are fully documented while others have only a label or
nothing at all.

This skill standardises the documentation so every class follows one
predictable pattern. The payoff is twofold:

1. **Human readability** -- A developer who has never touched VBA should be
   able to open any documented file and understand its purpose, public API,
   parameter meanings, return values, and error conditions.

2. **Machine extractability** -- A future parser will walk every `.cls` file,
   pull out the structured tags, and emit HTML reference pages grouped by
   class and section. The tag grammar must therefore be strict, consistent,
   and unambiguous so a simple regex- or line-based extractor can handle it.

---

## Trigger Keywords

Invoke this skill when the request matches any of:

- "document", "documentation", "doc comments", "add docs"
- "roxygen", "JSDoc", "HTML docs", "API reference"
- "make this class understandable", "explain this class"
- "write headers for", "add documentation to"
- Any request to produce or improve doc-comment blocks in `.cls` files
- References to building automated documentation from VBA source

---

## Critical Workflow

### Phase 1: Context Loading (MANDATORY)

Before writing a single doc comment, load the project context so the
documentation accurately reflects the codebase.

**Step 1.1 -- Read project rules:**
```
Read: <workspace_root>/obt-skill/project-rules.md
```
Section 6 in particular defines the existing documentation standards. Everything
you write must be compatible with the project conventions.

**Step 1.2 -- Read the tag reference:**
```
Read: <skill_root>/references/tag-reference.md
```
This file is the **grammar specification**. It defines every allowed tag, its
syntax, where it may appear, and whether it is required or optional.

**Step 1.3 -- Read the documentation template:**
```
Read: <skill_root>/references/doc-template.md
```
This file shows the canonical layout for a fully documented implementation
class and its matching interface. Use it as your blueprint.

**Step 1.4 -- Read the target file(s):**
Read the `.cls` file the user wants documented. If the class has an interface
(`IClassName.cls`), read that too -- you need both to produce consistent
cross-references (`@label` in implementation, `@jump` in interface).

**Step 1.5 -- Read related files when helpful:**
If the class depends on other project classes (look for `Dim x As ISomeClass`
or `Set x = SomeClass.Create(...)` patterns), skim those interfaces so you
understand the types flowing through parameters and return values.

---

### Phase 2: Analysis

Build a mental model of the class before writing anything.

1. **Purpose** -- What problem does this class solve? Distil it into one
   sentence. This becomes the `@ModuleDescription` and the `@description`
   in the class header.

2. **Sections** -- Group methods and properties into logical sections. Keep
   existing `@section` groupings whenever they make sense; add new ones only
   for orphaned members.

3. **Public surface** -- List every `Public` member. Each one MUST get a
   full doc block. No exceptions.

4. **Private helpers** -- List every `Private` member. Document anything that
   contains non-trivial logic. Simple one-line delegations can have a
   minimal block (label + sub-title).

5. **Dependencies** -- Which classes does this one consume? These populate
   `@depends` at the class level and optionally on individual methods.

6. **Error paths** -- Which methods call `ThrowError` or `Err.Raise`? Each
   needs a `@throws` tag.

---

### Phase 3: Writing Documentation

Apply the grammar from `references/tag-reference.md` and the layout from
`references/doc-template.md`. The core principles follow.

#### 3.1 Every Public Member Gets a Full Doc Block

A "full doc block" means at minimum:

```vba
'@label:<unique-id>
'@sub-title <one-line imperative summary>
'@details
'<Paragraph explaining behaviour, edge cases, and why it exists.
'Write for a beginner: explain what the method does in plain English,
'what happens when inputs are invalid, and any side effects.>
'@param <name> <Type>. <Description ending with a period.>
'@return <Type>. <Description ending with a period.>
'@throws <ErrorType> <When condition>.
'@depends <ClassA>, <ClassB>
'@export
```

Use `@sub-title` for `Sub` and `Function` members. Use `@prop-title` for
`Property Get/Let/Set` members.

#### 3.2 Interface Files Mirror Implementation Docs

The interface (`IClassName.cls`) must document every member with the same
level of detail, but uses `@jump:<label>` instead of `@label:<label>` to
cross-reference the implementation. This lets the HTML generator link an
interface member straight to its implementation.

#### 3.3 The Class Header Block

Every `.cls` file gets a standardised header right after the VBA attributes
and `Option Explicit`:

```vba
'@PredeclaredId                           ' if applicable
'@Folder("<TopicFolder>")
'@ModuleDescription("<one-line summary>")
'@IgnoreModule <inspections>

'@class <ClassName>
'@description
'<2-5 sentences: what this class does, who consumes it, how it fits
'into the project architecture. A beginner should be able to read this
'and understand when and why they would use this class.>
'@depends <ClassA>, <ClassB>, <ClassC>
'@version <version or date>
'@author <author if known>
```

#### 3.4 Section Headers

Group related members with clear section markers:

```vba
'@section <SectionName>
'===============================================================================
```

If the section name alone is not self-explanatory, add a one-line
`'@description` right below the divider.

#### 3.5 Parameter Documentation

Every parameter gets its own `@param` tag. The format is strict so a parser
can split it mechanically:

```
'@param <paramName> <Type>. <Description ending with a period.>
```

Rules:
- `<paramName>` must match the VBA signature exactly (case-sensitive).
- `<Type>` is the VBA type (`String`, `Long`, `Worksheet`, `IDataSheet` ...).
- For `Optional` parameters, state the default value in the description:
  `'@param colName Optional String. Column header to retrieve. Defaults to "__all__".`
- For `ByRef` parameters that are modified, say so:
  `'@param outCount ByRef Long. Populated with the row count on return.`

#### 3.6 Return Documentation

```
'@return <Type>. <Description ending with a period.>
```

If the method can return `Nothing` under certain conditions, state when in
`@details` and repeat it briefly in `@return`:
`'@return IDataSheet. The backing data store, or Nothing when uninitialised.`

#### 3.7 Error Documentation

```
'@throws <ProjectError.ErrorName> When <condition>.
```

Every call to `ThrowError` or `Err.Raise` in the method body should have a
matching `@throws` tag. This tells the caller what to guard against.

#### 3.8 Remarks and Notes

- `@remarks` -- implementation-level observations for maintainers (not part
  of the API contract).
- `@note` -- important caveats that affect callers.

```vba
'@remarks The sort relies on Range.Sort which is locale-sensitive.
'Test on both Windows and macOS.
'@note This method mutates the worksheet in place. There is no undo.
```

#### 3.9 Inline Comments vs Doc Tags

Doc tags live in the block ABOVE the `Sub`/`Function`/`Property` line. They
answer: "What does this do? What goes in? What comes out?"

Inline comments live INSIDE the method body. They answer: "Why is this line
here? What is the algorithm doing?"

Both matter, but this skill focuses on the doc-tag layer. Preserve existing
inline comments as-is unless they are factually wrong.

---

### Phase 4: Output Rules

These rules are non-negotiable because they come from the project's own
coding standards.

1. **Return the FULL file.** Never snippets, never diffs. Even if only
   documentation changes, output the entire `.cls` content.

2. **Preserve all existing code.** Do not change logic, variable names,
   indentation, or formatting. You are adding or improving documentation
   only.

3. **Keep correct existing tags.** A `@details` block is "correct" when it
   contains 2-5 sentences covering behaviour, edge cases, and context; is
   written so a beginner can understand it; and accurately matches the code.
   If an existing block meets those criteria, keep it. Improve blocks that
   are too brief (< 2 sentences), inaccurate, or missing required tags.

4. **Normalise tag variants across the entire file.** Because the automated
   parser expects strict consistency, convert ALL non-canonical tags in the
   file -- not just the ones on the method you are currently documenting:
   - `@params` (plural block) to individual `@param` lines
   - `@returned` or `@returns` to `@return`
   - `@pram` (typo) to `@param`
   - `@fun-title` to `@sub-title`
   - `@remark` (singular) to `@remarks`
   The tag reference defines every canonical form and its deprecated variants.

5. **Run unix2dos** on every file you produce:
   ```bash
   unix2dos <filename.cls>
   ```

6. **No emojis in code or comments. No markdown inside code files.**

7. **Update `.obt/tracking.md`** if this work is part of a tracked task.

---

### Phase 5: Verification Checklist

After writing, verify each of these before delivering:

- [ ] Every `Public` member has a full doc block
- [ ] Every `@param` matches a parameter in the VBA signature exactly (name is case-sensitive)
- [ ] `@param` tags are listed in the same order as the VBA signature
- [ ] Every function/property-get that returns a value has `@return`
- [ ] Every method that calls `ThrowError` or `Err.Raise` has `@throws`
- [ ] `@section` headers group related members logically
- [ ] Class header has `@class`, `@description`, `@depends`
- [ ] Interface `@jump` tags match implementation `@label` tags 1:1
- [ ] No deprecated tag variants remain (`@params`, `@pram`, `@fun-title`, etc.)
- [ ] All descriptions end with a period
- [ ] Tags appear in the recommended order (see `tag-reference.md` Section 1)
- [ ] File is complete (beginning to end, not a snippet)
- [ ] `unix2dos` has been run

---

## Tag Quick Reference

| Tag | Placement | Required | Purpose |
|-----|-----------|----------|---------|
| `@class` | Class header | Yes | Names the documented class |
| `@description` | Header / section | Yes (header) | Multi-line overview |
| `@section` | Before member groups | Yes | Logical grouping divider |
| `@label:<id>` | Implementation members | Yes | Unique anchor for cross-refs |
| `@jump:<id>` | Interface members | Yes | Points to implementation `@label` |
| `@sub-title` | Sub / Function | Yes | One-line summary |
| `@prop-title` | Property Get/Let/Set | Yes | One-line summary |
| `@details` | Any member | Yes (public) | Behaviour explanation |
| `@param` | Members with params | Yes | One tag per parameter |
| `@return` | Functions / Prop Get | Yes | Return value description |
| `@throws` | Error-raising members | When applicable | Error condition |
| `@depends` | Header / members | When applicable | Class dependencies |
| `@export` | Public factory methods | When applicable | Marks public API |
| `@remarks` | Any member | Optional | Maintainer notes |
| `@note` | Any member | Optional | Caller caveats |
| `@version` | Class header | Optional | Version or date |
| `@author` | Class header | Optional | Authorship |
| `@todo` | Any location | Optional | Future work marker |
| `@PredeclaredId` | Class header | When applicable | Enables predeclared factory pattern |
| `@Interface` | Interface header | Yes (interfaces) | Marks interface contracts |
| `@Folder` | Class header | Yes | Organises class in Rubberduck Code Explorer |
| `@ModuleDescription` | Class header | Yes | Short description for Rubberduck UI |
| `@IgnoreModule` | Class header | When applicable | Suppresses module-level inspections |
| `@Ignore` | Above target line | When applicable | Suppresses a single inspection |

For the complete specification with syntax rules and examples, see
`references/tag-reference.md`.

---

## When in Doubt

1. **Over-document** rather than under-document. A beginner needs more
   context, and a parser needs every tag present.
2. If you are unsure what a method does, add `@todo Clarify behaviour with
   maintainer` rather than guessing.
3. **Never "improve" logic** while documenting. Code changes require the
   `obt` skill, not this one.
4. Ask the user if a section grouping or description feels unclear.

---

## Version History

- **1.0** (2026-02-09): Initial skill creation

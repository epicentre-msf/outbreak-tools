# Running the automated test suite on macOS

This document explains, in full, how the OutbreakTools VBA test suite is run
**headlessly on macOS** — driven from the command line, with the code imported
from `src/` automatically, and the results collected as a file. No manual
upload into Excel, no clicking.

It covers the architecture, the exact command, the one‑time setup, the macOS
sandbox model (and why it behaves the way it does), how to read the results,
how to add new test suites, and a troubleshooting catalogue of every failure
mode we hit while building this.

> **Platforms.** Only the macOS path is implemented here. Windows parity (a
> trigger‑file watcher) is a separate, later effort. Everything below is macOS.

> **Working area.** The driver workbook and every per-run file live in an
> untracked, gitignored working area — not in the repo. Its location defaults to
> `.test-runner/` but is a local setting: pass `--home=<dir>` to run-tests.R or set the
> `OBT_TEST_HOME` env var. Paths below show the `.test-runner` default; read
> `.test-runner/tests/staging` as `<home>/tests/staging`.

---

## Table of contents

1. [The one command](#1-the-one-command)
2. [What actually happens (the chain)](#2-what-actually-happens-the-chain)
3. [The moving parts](#3-the-moving-parts)
4. [One‑time setup (do this once per machine/workbook)](#4-one-time-setup)
5. [The macOS sandbox story (why the folder grant matters)](#5-the-macos-sandbox-story)
6. [The registry: single source of truth](#6-the-registry-single-source-of-truth)
7. [The run directory (how sources reach Excel)](#7-the-run-directory)
8. [Reading the results](#8-reading-the-results)
9. [Adding a new test suite](#9-adding-a-new-test-suite)
10. [Tracked vs untracked (what lives where)](#10-tracked-vs-untracked)
11. [Troubleshooting catalogue](#11-troubleshooting-catalogue)
12. [Appendix: file/really-quick reference](#12-appendix-quick-reference)

---

## 1. The one command

From the repository root:

```bash
Rscript scripts/tests/run-tests.R --build
```

- `--build` regenerates the test manifest from the registry and rebuilds the
  workbook's Codes tables + `ModulesForTesting` before running. **Use it on the
  first run and whenever `test-registry.yml` changes.** Without `--build`, the
  run reuses whatever tables are already in the workbook.

The command:

- exits **0** when every test passed,
- exits **1** on any failure, an incomplete run, or **zero tests executed**
  (an all‑zero run is treated as a failure, not a pass),
- prints a per‑module summary and, on failure, the diagnostics log.

That is the whole user‑facing surface. Everything below is how it works and how
to set it up the first time.

---

## 2. What actually happens (the chain)

`run-tests.R` is the orchestrator (all portable logic lives here). The
macOS‑specific trigger, `scripts/tests/macos/run-tests.applescript`, is a thin
shim that does only what Apple Events can do: drive Excel.

```
Rscript run-tests.R --build
│
├─ 1. (--build) Rscript build-registry.R
│        parse src/tests/test-registry.yml
│        → src/tests/.generated/code-tables.tsv          (folder, tag, component, interface)
│        → src/tests/.generated/modules-for-testing.txt  (flat list of test modules)
│
├─ 2. Assemble the run dir  .test-runner/tests/staging/run/
│        unit_tests_run.xlsb          (copy of .test-runner/PartialTests.xlsb — original untouched)
│        classes/<folder>/…           (registered class sources, from src/ then .test-runner fallback)
│        tests/modules/<folder>/…      (registered test modules)
│        .generated/…                 (the manifest, copied in)
│        bootstrap/OBTImport.bas       (harness sources for the refresh step)
│        bootstrap/OBTHeadless.bas
│
├─ 3. Quit any running Excel, wait until the process is gone (pgrep)
│
├─ 4. osascript run-tests.applescript <copy> build
│        open workbook … read only false          (read/WRITE — the run mutates it)
│        run VB macro "…!OBTRefreshHarness"        (re-import OBTImport/OBTHeadless from run dir)
│        run VB macro "…!OBTBuildCodeTables"       (only if --build; rebuild Codes tables + ModulesForTesting)
│        run VB macro "…!OBTSilentImport"          (Development.ImportAll — pulls the probe/suite code)
│        run VB macro "…!OBTRunAllTests"           (run every module, serialize testsOutputs → CSV, Save)
│        quit saving no                            (AppleScript quits Excel; VBA already Saved)
│
└─ 5. read .test-runner/tests/staging/run/test-results.csv
         summarise: total, passed, failed, per-module failures
         copy latest → .test-runner/tests/staging/test-results.csv
```

**Order matters.** `OBTBuildCodeTables` runs *before* `OBTSilentImport`, because
the silent import reads the freshly‑rebuilt Codes tables to decide what to pull
from disk. And `OBTRefreshHarness` runs *first*, so the current harness code is
loaded before any of the steps that use it.

---

## 3. The moving parts

### Host side (tracked, in the repo)

| File | Role |
|---|---|
| `scripts/tests/run-tests.R` | Orchestrator. Builds the run dir, quits/relaunches Excel via the trigger, reads and summarises results. Exit code drives CI. |
| `scripts/tests/build-registry.R` | Parses `test-registry.yml` → the two `.generated` manifest files. YAML is parsed in R (never in VBA). |
| `scripts/tests/macos/run-tests.applescript` | Thin Apple Events trigger: open read/write, run the `OBT*` macros in order, quit Excel. |
| `src/tests/test-registry.yml` | **Single source of truth** for which components/tests exist and where they live. |

### Workbook side (VBA — the harness)

Tracked in `src/tests/modules/rubberduck/`:

| Module | Entry point(s) | Role |
|---|---|---|
| `OBTBootstrap.bas` | `OBTRefreshHarness` | Re‑imports `OBTImport`/`OBTHeadless` from the run dir at the start of each run, so the harness can be iterated **without a manual VBE re‑import**. The one module that still needs a manual re‑import if *it* changes. |
| `OBTImport.bas` | `OBTBuildCodeTables`, `OBTSilentImport` | Wraps the `Development` manager. `OBTBuildCodeTables` rebuilds the Codes import tables + the `ModulesForTesting` table from the manifest; `OBTSilentImport` runs `Development.ImportAll` with all prompts/alerts off. Self‑heals the `Codes`/`testsOutputs` sheets. Writes `obt-import.log`. |
| `OBTHeadless.bas` | `OBTRunAllTests`, `OBTHeadlessActive` | Discovers `'@TestMethod` procs via the VBProject, runs each module's lifecycle, serialises `testsOutputs` → `test-results.csv`. Self‑contained (no ribbon, no `TestHelpers`). |
| `OBTGrantAccess.bas` | `OBTGrantAccess` | **One‑time, interactive.** Pops a macOS folder picker so you grant Excel persistent sandbox access to the repo tree. Not part of the automated loop. |

`Development.cls` (the import manager) and its dependency tree are also in the
workbook — see [setup](#4-one-time-setup).

### The workbook

`.test-runner/PartialTests.xlsb` — the driver workbook. It carries the harness + the
`Development` manager, plus two worksheets: **`Codes`** (import tables +
`ModulesForTesting`) and **`testsOutputs`** (rendered results). It is an
**untracked binary artefact** (under `.test-runner/`, which is git‑ignored). The
runner copies it per run and never mutates the original.

---

## 4. One‑time setup

Do this once per machine (and once per fresh workbook). After it's done, every
future run is just `Rscript scripts/tests/run-tests.R --build`.

### 4.1 Import the harness + `Development` into the workbook

Open `.test-runner/PartialTests.xlsb` and, in the VBE (**File → Import File…**),
import these components. They form the compile closure needed for `Development`
+ the harness to compile:

```
Development manager + dependencies
  BetterArray.cls        (src/classes/general)
  Checking.cls           (src/classes/general)   ← also home of the ProjectError enum
  CheckingOutput.cls     (src/classes/general)
  HiddenNames.cls        (src/classes/general)
  TranslationObject.cls  (src/classes/general)
  Passwords.cls          (src/classes/general)
  Development.cls        (src/classes/dev)

Harness
  CustomTest.cls         (src/classes/rubberduck)
  OBTHeadless.bas        (src/tests/modules/rubberduck)
  OBTImport.bas          (src/tests/modules/rubberduck)
  OBTBootstrap.bas       (src/tests/modules/rubberduck)
  OBTGrantAccess.bas     (src/tests/modules/rubberduck)
```

> **You do NOT import the probe/test classes by hand.** Importing the code under
> test is exactly what the loop automates (`OBTSilentImport`). Only the harness
> + `Development` are imported manually, and only once.

You do **not** need to create the `Codes` or `testsOutputs` sheets by hand —
`OBTImport` self‑heals them (`EnsureSheet` creates either if missing; a blank
`testsOutputs` is fine, `CheckingOutput` initialises its own named ranges).

### 4.2 Turn on VBA project access

**Excel → Preferences → Security → “Trust access to the VBA project object
model” = ON.** Required for `@TestMethod` discovery and for `OBTRefreshHarness`
to re‑import modules.

### 4.3 Set VBE error trapping

**VBE → Preferences → General → Error Trapping = “Break on Unhandled Errors”.**
The harness leans on `On Error Resume Next`; “Break on All Errors” halts a
headless run.

### 4.4 Grant Excel folder access (the important one)

Run **`OBTGrantAccess` once, by hand**:

1. In the VBE, put the cursor inside `OBTGrantAccess` and press **F5**.
2. In the folder picker, select the **`outbreak-tools` repo root** and confirm.

This is the step that stops the “grant access to files” prompts. See
[§5](#5-the-macos-sandbox-story) for why it's necessary and why Full Disk Access
is **not** a substitute. One pick of the repo root covers every current and
future test file, and the grant **persists across Excel quit→relaunch**.

### 4.5 Close Excel

Fully quit Excel before running. The runner drives a quit→fresh‑launch cycle,
and `run VB macro` fails against an already‑open interactive Excel.

---

## 5. The macOS sandbox story

This is the part that surprises people, so it's worth stating plainly.

**Excel for Mac is a sandboxed application.** Full Disk Access is a *TCC privacy*
setting; it does **not** lift the *App Sandbox*, which is what governs VBA's
programmatic file access (`Open`, `Dir`, `VBComponents.Import`). So:

- Granting Excel **Full Disk Access does nothing** for our file reads — you'll
  still be prompted.
- A sandboxed app gets **persistent, whole‑subtree** access to a folder only
  when a user **picks it through a file/folder dialog** (a macOS
  *security‑scoped bookmark*). That is exactly what `OBTGrantAccess` (and, in
  normal OutbreakTools use, `clickDevFolder` → `OSFiles.LoadFolder`) does.

Two consequences baked into the design:

- **`OBTGrantAccess` grants the repo tree once.** Verified: the grant survives
  Excel's quit→relaunch, so subsequent headless runs are silent, for any number
  of files.
- **The run directory is a stable path** (`.test-runner/tests/staging/run`), not a
  fresh `run-<timestamp>` each time. macOS grants file access *per path*; a new
  path every run is what caused repeated prompts.

If you ever *do* get prompts, it means either `OBTGrantAccess` wasn't run, or a
path outside the granted tree is being touched.

---

## 6. The registry: single source of truth

`src/tests/test-registry.yml` declares every suite. A suite is a `folder` plus
its classes, general modules, test‑only fixtures, and test modules. Example:

```yaml
suites:
  - folder: sections
    classes:
      - name: VarWriter
        interface: true          # also imports IVarWriter.cls if present
      - name: SectionBuilder
        interface: true
    tests:
      - module: TestVarWriter
        covers: [VarWriter]
      - module: TestSectionBuilder
        covers: [SectionBuilder]
```

For a suite `folder: F`, components resolve to:

| Registry key | Disk path | Codes‑table tag |
|---|---|---|
| `classes[].name` | `src/classes/F/<name>.cls` (+ `I<name>.cls` if `interface: true`) | `general classes` |
| `modules[]` | `src/modules/F/<name>.bas` | `general modules` |
| `fixtures[].name` | `src/tests/classes/F/<name>.cls` | `tests classes` |
| `tests[].module` | `src/tests/modules/F/<module>.bas` | `tests modules` |

`build-registry.R` flattens all of this into:

- `code-tables.tsv` — one row per component: `folder <TAB> tag <TAB> component <TAB> interface`.
  `OBTBuildCodeTables` groups these by `(folder, tag)` and builds one Codes
  import table per group (with the folder written into the cell above the
  header, which `Development` reads as the subfolder).
- `modules-for-testing.txt` — the flat union of every test module.
  `OBTBuildCodeTables` rebuilds the `ModulesForTesting` table from it, and
  `OBTRunAllTests` iterates that table to decide what to run.

**A test module runs only once it appears in the registry AND `--build` has
repopulated the workbook.** That is by design — add suites progressively.

---

## 7. The run directory

Excel never reads directly from `src/`. Instead `run-tests.R` assembles a
self‑contained **run directory** next to the workbook copy, and the workbook
reads everything relative to its own folder (`ThisWorkbook.Path`). Because that
folder is inside the granted repo tree, all reads/writes are prompt‑free.

For each registered `(folder, tag)` the runner copies the folder's sources into
the run dir, resolving **`src/` first, then the `.test-runner` staging** as a fallback.
That fallback is what lets the local `draft` probe (whose fixtures are
intentionally untracked) run alongside real suites sourced from `src/`.

Layout of `.test-runner/tests/staging/run/` during a run:

```
unit_tests_run.xlsb     the workbook copy (opened by AppleScript)
classes/<folder>/*.cls  general classes           → ClassesImplementation
modules/<folder>/*.bas  general modules           → ModulesCodes
tests/modules/<folder>/ test modules              → TestsCodes/modules
tests/classes/<folder>/ test-only fixtures        → TestsCodes/classes
.generated/*            the manifest (read by OBTBuildCodeTables)
bootstrap/*.bas         OBTImport + OBTHeadless    (read by OBTRefreshHarness)
test-results.csv        written by OBTRunAllTests
obt-import.log          written by OBTImport
```

`OBTImport` writes the three folder‑path named ranges (`ModulesCodes`,
`ClassesImplementation`, `TestsCodes`) into the `Codes` sheet **in code** at run
start, pointing at these run‑dir subfolders — so no Dev‑sheet wiring is needed
by hand.

---

## 8. Reading the results

`OBTRunAllTests` serialises the `testsOutputs` sheet to
`.test-runner/tests/staging/run/test-results.csv`:

```
module,title,type,label,message
"TestDefaultNameIsWorld","TestDefaultNameIsWorld","success","Success: …","Success: …"
…
"__summary__",,"","total=16 failures=0","modulesRun=2 testsFound=14 vbe=ok(19) name=listobject:B3:B4"
```

- `type` is `success` or `error` (the sheet stores status glyphs; the serializer
  normalises them). Empty `type` rows are not counted.
- The final `__summary__` row is diagnostics, not a test:
  - `total` / `failures` — assertion counts.
  - `modulesRun` — how many modules the runner iterated. **`0` means
    `ModulesForTesting` was empty or missing.**
  - `testsFound` — how many `@TestMethod` procs were discovered.
  - `vbe` — `ok(<n>)` means the VBProject was reachable (`n` components); an
    error code means “Trust access to the VBA project object model” is off.
  - `name` — how `ModulesForTesting` was resolved (`listobject:<addr>`), or
    `not-found`.

`run-tests.R` prints the summary line, the pass/fail totals, and — on failure —
the per‑module failing rows and the contents of `obt-import.log`. On success it
copies the CSV to `.test-runner/tests/staging/test-results.csv`.

---

## 9. Adding a new test suite

1. Write the class(es) under `src/classes/<folder>/` and the test module(s)
   under `src/tests/modules/<folder>/`, following the `CustomTest` pattern
   (`CustomTest.Create` + `@TestMethod` + `Assert.AreEqual`/`LogFailure` — never
   `Err.Raise` in a test).
2. Register them in `src/tests/test-registry.yml`.
3. Make sure the workbook can compile the new code: its **dependency closure**
   must be present in the workbook. `OBTSilentImport` imports the registered
   components, but any *other* classes they reference must already be imported
   (once) as part of setup — otherwise the project won't compile and the run
   fails. (This is why the throwaway probe uses self‑contained classes.)
4. Run `Rscript scripts/tests/run-tests.R --build`.

> Iterating on the **harness** (`OBTImport`/`OBTHeadless`) needs no manual
> re‑import — `OBTRefreshHarness` reloads them from the run dir each run. Only
> `OBTBootstrap` itself, if changed, needs a one‑time manual re‑import.

---

## 10. Tracked vs untracked

| Lives in `src/` (tracked) | Lives in `.test-runner/` (untracked, git‑ignored) |
|---|---|
| `scripts/tests/*` (orchestrator, trigger, registry builder) | `PartialTests.xlsb` (the driver workbook — a binary artefact) |
| `src/tests/test-registry.yml` | `tests/staging/run/` (the per‑run working dir) |
| `src/tests/modules/rubberduck/OBT*.bas` (the harness) | `tests/staging/{classes,tests}/draft/…` (the local **probe** fixtures) |
| Real project classes + their real test suites | `tests/staging/bootstrap/*` (staging copies) |

**The `draft` probe classes are test‑only smoke‑test fixtures and are
deliberately NOT committed to `src/`.** They exercise the loop itself (import →
run across more than one class) and live only in `.test-runner`. Real project suites
go in `src/` and are the point of the loop.

---

## 11. Troubleshooting catalogue

Every one of these was hit for real while building this loop.

| Symptom | Cause | Fix |
|---|---|---|
| **`0 success / 0 failure`, `modulesRun=0 name=not-found`** | `OBTBuildCodeTables` didn't build `ModulesForTesting` (errored earlier), or `Codes`/`testsOutputs` missing. | Read `obt-import.log` for the failing step. The runner already treats this as a failure (not a false green). |
| **AppleScript `Parameter error (-50)`** | `run VB macro` against an already‑open interactive Excel. | Fully quit Excel before running; the orchestrator already quits + waits on `pgrep`. |
| **AppleScript `AppleEvent timed out (-1712)`** | A macro hung — usually a modal dialog (a compile‑error dialog, or a file‑access prompt) blocking the Apple Event. | Ensure the project compiles and that `OBTGrantAccess` was run so no file‑access modal fires. |
| **Repeated “grant access to files” prompts** | Excel is sandboxed; each new path prompts. FDA does not help. | Run `OBTGrantAccess` once (pick the repo root). The run dir is a stable path, so one grant covers all runs. |
| **`error 91` at the first sheet operation** | A local variable named case‑insensitively the same as a module symbol (e.g. `codeSheet` vs `Const CODESHEET`) shadows it, so the wrong value is passed. | Give module constants underscore names (`CODE_SHEET`); never name a variable like another symbol in the file. |
| **Harness change didn't take effect** | The loop only re‑imports the **probe/suite** code, not the harness — unless `OBTRefreshHarness` ran. | Confirm the AppleScript calls `OBTRefreshHarness` first; confirm `run-tests.R` staged the harness into `run/bootstrap/`. Only `OBTBootstrap` needs a manual re‑import. |
| **A whole suite fails to compile after import** | Its dependency closure isn't in the workbook. | Import the missing dependency classes once (setup), or keep suites self‑contained. |
| **Can't tell which VBA actually ran** | `strings` on `vbaProject.bin` is useless — VBA source is compressed there. | Decompile with `oletools` (`olevba` / `VBA_Parser`) to read the real module source. |
| **A code‑module *write* stalls under automation** | `CodeModule.InsertLines` (e.g. `CheckingOutput`'s worksheet‑change handler injection) stalls under Apple Events. | Already guarded: `CustomTest.PrintResults` skips the injection when `OBTHeadlessActive` is True. (`VBComponents.Import` — component import — works fine; it's only in‑place code writes that stall.) |

---

## 12. Appendix: quick reference

**Run:**
```bash
Rscript scripts/tests/run-tests.R --build     # first run / after registry change
Rscript scripts/tests/run-tests.R             # reuse existing Codes tables
Rscript scripts/tests/run-tests.R --home=DIR  # workbook + staging live under DIR (default .test-runner)
```

**Macro order (inside the workbook, per run):**
`OBTRefreshHarness` → [`OBTBuildCodeTables` if `--build`] → `OBTSilentImport` → `OBTRunAllTests`

**Key paths:**
- Registry: `src/tests/test-registry.yml`
- Manifest: `src/tests/.generated/{code-tables.tsv, modules-for-testing.txt}`
- Harness: `src/tests/modules/rubberduck/OBT*.bas`
- Workbook: `.test-runner/PartialTests.xlsb`
- Run dir: `.test-runner/tests/staging/run/`
- Results: `.test-runner/tests/staging/test-results.csv` (+ `run/test-results.csv`, `run/obt-import.log`)

**One‑time setup checklist:**
- [ ] Import the harness + `Development` + its deps into `PartialTests.xlsb`
- [ ] Trust access to the VBA project object model = ON
- [ ] VBE Error Trapping = “Break on Unhandled Errors”
- [ ] Run `OBTGrantAccess` once, pick the repo root
- [ ] Quit Excel before running

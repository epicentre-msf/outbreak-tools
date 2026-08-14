# The headless workflow

Two things this project could previously only do by hand: fill a setup workbook
from another setup, and generate a linelist. Both run inside Excel, both are
driven from a ribbon, and both need a designer workbook whose baked code is
whatever somebody last pasted into it.

This folder drives them from code instead, so a run proves the source in `src/`
rather than a binary nobody can diff. It is also the layer an R package is meant
to sit on: every entry point takes its paths as parameters, answers a **string**
outcome, and raises nothing.

```
"OK"                          the run finished
"ERROR <number>: <text>"      it did not, and this is why
```

A caller reads that string. It never has to read a dialog.

## The three pieces

| Where | What |
|---|---|
| `src/modules/headless/HeadlessBuild.bas` | the two entry points, imported into the driver workbook by the registry |
| `scripts/headless/vba/OBTSetupImportHeadless.bas` | injected into the setup being filled, run there, removed again |
| `scripts/headless/merge-form-code.R` | puts the current `FormLogic*` code into the exported `.frm` files |
| `scripts/headless/build-linelist.R` | generates a linelist from a filled setup, and runs no test |
| `scripts/headless/macos/build-linelist.applescript` | the trigger behind it, thin by design |

## Step one — fill a setup

```vb
outcome = HeadlessBuild.ImportSetupFromWorkbook(targetPath, sourcePath, modulePath)
```

`SetupImport` resolves its destination as `ThisWorkbook`, in seven places. Called
from a driver it would import the source **into the driver**. So the code that
calls it has to live inside the workbook being filled: `ImportSetupFromWorkbook`
reads `vba/OBTSetupImportHeadless.bas` off disk, imports it into the target's
VBProject, runs its entry point through `Application.Run`, and removes it again
before saving.

The target is opened read/write and saved. Pass a **copy** unless you mean to
change the original.

## Step two — merge the form code

```
Rscript scripts/headless/merge-form-code.R
```

A `.frm` exported from a workbook carries three parts: the `VERSION` line, the
`Begin … End` control tree, and the `Attribute VB_*` block, then the code. The
forms in `.mock/forms/designer` carry whatever code was in the workbook the day
they were exported; the code that *belongs* in them lives in
`src/modules/linelistform/FormLogic*.bas`. So a form imported straight from
`.mock` runs stale logic.

The script rewrites only the code and copies the first two parts byte for byte,
dropping the FormLogic module's own `Attribute` lines — a second `VB_Name` inside
a form module is what makes an import fail. Output lands in
`.test-runner/forms/merged`, eleven `.frm` files and their `.frx`.

`F_Progr` has no FormLogic module and is copied unchanged. `F_ShowHideLL` takes
`FormLogicShowHide`, which is the one pair whose names do not match; the mapping
is read off each module's `@ModuleDescription`, not off its file name.

## Step three — generate a linelist

```
Rscript scripts/headless/build-linelist.R --setup=<filled-setup.xlsb> \
        [--out=<dir>] [--name=<stem>] [--temppath=<ribbon.xlsb>] ...
```

which is one command that does the whole thing: merges the forms, stages the
source, drives Excel, and prints what the build recorded. Under it:

```vb
outcome = HeadlessBuild.BuildLinelistFromSetup( _
              designerPath, setupPath, sourceRoot, formsFolder, _
              outputFolder, outputName, options, grantRoot)
```

`grantRoot` is the one folder every other path sits under, and it is what the
sandbox is asked about. Left empty, the five paths are granted separately, which
is what `TestHeadlessLinelistBuild` still does — see **File access** below.

Three files land in `outputFolder`:

```
<outputName>.xlsb                 the linelist
<outputName>-generation.txt       the run log
<outputName>-designer.xlsb        the designer the run actually used
```

`options` is pipe-separated `key=value`, every key optional:

| key | meaning |
|---|---|
| `temppath` | the ribbon template. **Empty means the buttons build** — action buttons on the sheets, an Admin sheet, no ribbon |
| `geopath` | the geobase to import. Empty is valid and common |
| `setuplang` | the language of the setup file — a **column name** of the setup's own translation table, e.g. `English` |
| `lllang` | the language the linelist is written in — a **code**, or a `CODE-Name` entry, e.g. `ENG` or `ENG-English` |
| `llpassword` | the open password of the saved file. Empty opens on a double-click |

An unknown key is **reported**, not ignored: a misspelled key reads exactly like
an option nobody passed, and that difference is a build pointed somewhere else.

What the run did is read back through six properties — `LastReport`,
`LastSheetCount`, `LastVariableCount`, `LastComponentCount`, `LastLinelistPath`,
`LastLogPath`. A caller holding `"OK"` still knows nothing about the size of what
was built, and an empty linelist reports `"OK"` as readily as a full one.

### What the script does that the entry point does not

Generating used to happen only as a side effect of running
`TestHeadlessLinelistBuild`, so making a file took a test run, and the paths it
generated from were constants compiled into a test module — absolute, and true
on one machine. `build-linelist.R` is that build lifted out. Every path is an
argument, nothing is asserted, and the exit code says whether a linelist was
written.

| | |
|---|---|
| `--setup` | required: the filled setup to generate from |
| `--designer` `--out` `--name` | default to the mock designer, `<home>/build`, `linelist` |
| `--forms` | a folder of already-merged forms to copy in. Only meaningful with `--no-merge` |
| `--temppath` `--geopath` `--setuplang` `--lllang` `--llpassword` | the option keys above, passed through |
| `--no-merge` | skip step two. Skipping it is how a linelist ends up with the right buttons wired to last week's handlers |
| `--home` | the untracked working area, same one the test harness uses |
| `--obt-home` | **the granted root** — everything Excel touches is staged under it. Falls back to the `OBT_HOME` environment variable (an `.Renviron` line is the deployed way to set it), then to `<repo>/.headless-runner` |

It reads `src/tests/test-registry.yml` for the list of classes and modules to
load into the driver, and **drops every test row** before staging — the driver
gets the build closure and no test module at all. The registry is shared because
the closure is real and already written down once: `LinelistSpecs`, `Linelist`,
`CodeTransfer`, `TemporaryRepos`, `ApplicationState`, `InitTransfer` and
everything they type. A second list of it would drift, and a drifted list does
not fail — it delivers last month's code in a file that looks right. Narrowing
the registry for a probe only ever comments test rows, so a narrowed registry
still builds.

`HeadlessBuild.LastBuildSummary` is what the trigger reads the run back through.
The six accessors beside it are `Property Get`, and **`Application.Run` cannot
reach a `Property Get`** — a caller outside the process could read the outcome
string and nothing else, so `"OK"` arrived with no idea whether the linelist
held three sheets or none.

### The two languages are two different vocabularies

`setuplang` and `lllang` look interchangeable and are not, and getting them
confused is silent.

`setuplang` lands in `RNG_LangSetup`. It is a **column name of the setup's own
translation table** — `English` — and it translates the dictionary, the choices
and the analyses.

`lllang` lands in `RNG_LLForm`, and only its **code prefix** survives:
`InitTransfer` splits the value on the dash and writes the prefix as
`RNG_LLLanguageCode`. That code is the column name the four `LinelistTranslation`
tables are keyed on, and those columns are `ENG`, `FRA`, `SPA`, `POR`, `ARA`. So
`ENG-English` and a bare `ENG` both work, and `English` does not.

A code that names no column translates **nothing**, with nothing said: the
lookup falls back to returning the tag it was given, so the delivered linelist
comes out with worksheets called `LLSHEET_Analysis` and every button and message
reading as its own tag. `HeadlessBuild.ResolveInterfaceLanguage` is what stops
that now — it reads the codes off the designer's own `T_TradLLMsg` header row and
tries the caller's option, then the designer's own entry, then the first code the
table offers, reporting which one it settled on in the run narrative.

Leave both empty and the setup language is read off the setup, the way loading a
setup by hand fills it, and the linelist language stays whatever the designer
holds.

### Why a designer file is still copied

"No designer" means no designer *machinery* — no ribbon press, no
`DesignerEntry`, no form, no dialog. It does not mean no designer *worksheets*.
The build reads eight of them: `Main` carries the entries, `__formatter` the
design, `__pass` the passwords, `__formula` the tokens, `LinelistTranslation`
the five translation tables, `Geo` the geobase, `DesignerTranslation` the labels
and `__check` the run log. Those are data, and no sensible amount of code builds
them from nothing.

So a designer file is copied and its **code is thrown away**. Every non-document
component is removed from the copy, then the source is imported: the `.cls` files
of nine source folders, the `.bas` files of one, and the ten merged `.frm`.
`Linelist.TransferAllCode` then exports from the copy, so the delivered linelist
carries the source in the repository.

The copy's VBProject never has to *compile* — `CodeTransfer` exports components,
and an export reads text. That is why loading all of that into a workbook holding
a stale designer is safe here and would not be safe in a workbook that had to
run.

### Which folders, and why only those

| imported | skipped |
|---|---|
| `src/classes/{analyses,dataio,dictionary,general,geo,graphs,linelist,sections,showhide}` | `msetup`, `mastersetup`, `setup`, `designer`, `dev`, `rubberduck`, `stale`, `formulas` |
| `src/modules/{linelist}` | `linelistform`, `designer`, `dev`, `mastersetup`, `setup`, `headless` |

Those folders hold every component the transfer moves. The rest hold none — the
mastersetup disease classes have never been part of a linelist, and `formulas`
runs inside the *driver*, never inside the generated file.

`sections` is imported for **one class**. `SectionMap` is the only thing the
transfer takes out of it, and `EventsLinelistButtons` types `SectionMap`, so a
linelist built without the folder loses its project compile.

`src/modules/linelistform` is **not** imported, and that is the point. Its ten
`FormLogic*` modules are the code *behind the forms*: every one of them uses `Me`
and declares control event handlers, neither of which compiles in a standard
module. `merge-form-code.R` writes each one into the code module of the form it
belongs to, so the `.frm` files already carry that code where it is legal.
Importing the `.bas` files as well put four of them into the delivered linelist as
standard modules, and that alone cost the file its compile.

The narrow list is not tidiness. **Excel for Mac is sandboxed**: reading a folder
it has no security-scoped grant for pops a dialog, and a dialog in a headless run
is a hang. Walking the whole tree asked for those folders one prompt at a time.

The **strip is what makes the narrow list safe**. A component the transfer asks
for that no listed folder supplied is simply absent from the copy, so
`CodeTransfer` raises `ElementNotFound` naming it. Without the strip, the copy's
own stale component of that name would be exported and nothing would be said —
which is the failure mode worth designing against, because it produces a linelist
that looks right and behaves like last month.

## File access on macOS

Every path this workflow reads has to be inside a folder Excel holds a
security-scoped grant for. Full Disk Access does **not** provide one: the
sandbox wants bookmarks, and TCC permissions are a different thing entirely.

### One folder, one grant

Everything Excel touches is staged under one root, `OBT_HOME`, and the trigger
grants that one folder before its first read.

The direct reason is that **nothing granted the run dir at all**. `OBTImport`
and `OBTHeadless` make no grant call, `OBTRefreshHarness` reads
`run/bootstrap/*.bas` before anything else happens, and the only grant in the
chain sat three steps later naming five paths that did not include it. One root
granted at step 0 fixes that on its own.

Whether it gets to **zero** panels rather than fewer rests on two properties of
a folder grant that are **not established**: that it persists across an Excel
restart, and that it cascades to files created in the tree afterwards. Microsoft
documents `GrantAccessToMultipleFiles` as persistent; cascade has no
documentation. A probe meant to settle both was void — it derived its root from
`Environ$("HOME")`, which inside a sandboxed Excel answers the container home
where no grant is needed. See `.obt/gotchas/macos-sandbox-grant.md` before
citing either property.

The layout:

```
<OBT_HOME>/run/      the driver copy, bootstrap/, the filtered sources,
                     .generated/, build-report.txt, obt-import.log
<OBT_HOME>/src/      classes/** + modules/**, the tree the linelist is built
                     from (what the build calls sourceRoot)
<OBT_HOME>/forms/    the merged .frm files
<OBT_HOME>/in/       designer, setup, ribbon template, geobase
<OBT_HOME>/out/      the linelist, its log, the designer used, OBTApp_/
```

`build-linelist.R` is not sandboxed, so it copies the operator's files **in**
beforehand and copies the finished linelist back **out** to `--out` at the end.
Excel is handed the staged copies and never the originals, so it sees no path
outside the root wherever the operator keeps their setup.

The trigger's **step 0** is `OBTGrantRoot`, and it runs before every other
macro because every other macro reads a file. It lives in `OBTBootstrap` — the
module baked into the driver — since a module that has to be read off disk
cannot be the one that grants access to the disk. `OBTRefreshHarness` runs
straight after it and reads `run/bootstrap/*.bas`; nothing granted that folder
before, which is why the harness modules used to prompt on every single run.

What the run reports is in `build-report.txt` as `grant=`, and
`build-linelist.R` prints it on every run. A grant that stopped taking is
otherwise invisible from outside: no exit code and no elapsed time can see a
dialog.

### The scratch folder still has to exist before the grant

`CodeTransfer` exports every component of the transfer to `<lldir>/OBTApp_` and
imports it straight back, so a build touches two files in there per component.
`TemporaryRepos` creates the folder, but not until the build is well past the
one moment a dialog could still be answered — so `BuildLinelistFromSetup`
creates `<outputFolder>/OBTApp_` itself first. Under one granted root that is
belt and braces rather than load-bearing, and it costs nothing.

### Dead ends, so nobody walks them again

- An AppleScript `choose folder` sent to Excel **from outside** (osascript) does
  not create Excel's persistent grant — the panel does not run in Excel's
  sandbox context. Only in-process mechanisms stick.
- Full Disk Access does nothing for this.
- `pkill` on Excel can cost the grant. The bookmark is written on normal
  termination, so `build-linelist.R` quits Excel through Apple Events.

On Windows the sandbox does not exist, `GrantAccessToMultipleFiles` is absent,
`OBTGrantRoot` answers "not available on this host", and `EnsureFileAccess`
silently answers False. All correct there.

## Running the suites

Both steps have a suite under `src/tests/headless/`, registered under
`folder: headless`:

- `TestHeadlessSetupFill` — fills `setup_dev.xlsb` from `generic-test-setup.xlsb`
  and counts the dictionary rows that landed.
- `TestHeadlessLinelistBuild` — generates from the filled setup, then **reopens
  the built file and counts the components of its project**. It is the only suite
  that runs `Linelist.Prepare` to completion, because the code transfer needs a
  source project carrying the ten forms and the driver workbook carries none.

Run `merge-form-code.R` before the harness. Without it the suite reports the
missing folder and names the command, rather than building a linelist from stale
form code. `build-linelist.R` runs the merge itself, so this applies to the test
harness alone.

The two commands do two jobs and neither does the other's:

| | |
|---|---|
| `Rscript scripts/tests/run-tests.R --build` | runs the suites, asserts, writes `test-results.csv` |
| `Rscript scripts/headless/build-linelist.R --setup=…` | writes a linelist, asserts nothing |

`TestHeadlessLinelistBuild` still generates its own file, because what it asserts
on is a file it watched being made. A build for a person to open comes from the
build script.

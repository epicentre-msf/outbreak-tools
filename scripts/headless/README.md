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
| `src/modules/headless/HeadlessBuild.bas` | the two entry points, imported into the driver workbook by the test registry |
| `scripts/headless/vba/OBTSetupImportHeadless.bas` | injected into the setup being filled, run there, removed again |
| `scripts/headless/merge-form-code.R` | puts the current `FormLogic*` code into the exported `.frm` files |

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

```vb
outcome = HeadlessBuild.BuildLinelistFromSetup( _
              designerPath, setupPath, sourceRoot, formsFolder, _
              outputFolder, outputName, options)
```

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
| `setuplang` | the language of the setup file |
| `lllang` | the language the linelist is written in |
| `llpassword` | the open password of the saved file. Empty opens on a double-click |

An unknown key is **reported**, not ignored: a misspelled key reads exactly like
an option nobody passed, and that difference is a build pointed somewhere else.

What the run did is read back through six properties — `LastReport`,
`LastSheetCount`, `LastVariableCount`, `LastComponentCount`, `LastLinelistPath`,
`LastLogPath`. A caller holding `"OK"` still knows nothing about the size of what
was built, and an empty linelist reports `"OK"` as readily as a full one.

### Why a designer file is still copied

"No designer" means no designer *machinery* — no ribbon press, no
`DesignerEntry`, no form, no dialog. It does not mean no designer *worksheets*.
The build reads eight of them: `Main` carries the entries, `__formatter` the
design, `__pass` the passwords, `__formula` the tokens, `LinelistTranslation`
the five translation tables, `Geo` the geobase, `DesignerTranslation` the labels
and `__check` the run log. Those are data, and no sensible amount of code builds
them from nothing.

So a designer file is copied and its **code is thrown away**. Every non-document
component is removed from the copy, then 84 are imported: the `.cls` files of ten
source folders, the `.bas` files of two, and the eleven merged `.frm`.
`Linelist.TransferAllCode` then exports from the copy, so the delivered linelist
carries the source in the repository.

The copy's VBProject never has to *compile* — `CodeTransfer` exports components,
and an export reads text. That is why loading all of that into a workbook holding
a stale designer is safe here and would not be safe in a workbook that had to
run.

### Which folders, and why only those

| imported | skipped |
|---|---|
| `src/classes/{analyses,dataio,dictionary,general,geo,graphs,linelist,showhide}` | `msetup`, `mastersetup`, `setup`, `designer`, `dev`, `rubberduck`, `stale`, `formulas`, `sections` |
| `src/modules/{linelist,linelistform}` | `designer`, `dev`, `mastersetup`, `setup`, `headless` |

Ten folders hold all 55 components the transfer moves. The rest hold none — the
mastersetup disease classes have never been part of a linelist, and `formulas`
and `sections` run inside the *driver*, never inside the generated file.

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
security-scoped grant for. Full Disk Access does **not** provide one; only a
folder pick through a dialog does.

The run asks for it itself. Every headless entry point calls
`HeadlessBuild.EnsureFileAccess` — a thin wrapper over
`Application.GrantAccessToMultipleFiles`, Excel's own API for sandbox grants —
**before its first file read**, passing the repo root. The first run on a
machine shows one consolidated dialog; click Grant and the run proceeds. The
grant persists, so every later run is silent.

Two dead ends, so nobody walks them again:

- An AppleScript `choose folder` sent to Excel **from outside** (osascript)
  does not create Excel's persistent grant — the panel does not run in Excel's
  sandbox context. Only in-process mechanisms stick: `MacScript` from the VBE
  (what `OBTGrantAccess` does) or `GrantAccessToMultipleFiles` (what the code
  does now).
- Full Disk Access does nothing for this: the sandbox wants security-scoped
  bookmarks, not TCC permissions.

There is no way to drop the machinery on macOS — the sandbox is the host's, not
this project's. What the code does instead is ask **once, for one root, in one
dialog**. On Windows the sandbox does not exist, `GrantAccessToMultipleFiles`
is absent, and `EnsureFileAccess` silently answers False, which is correct
there.

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
form code.

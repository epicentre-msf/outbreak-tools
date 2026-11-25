---
name: headless-failure
description: Investigate a failure in a headless linelist build, a setup conversion, a verify case or a unit test run. Where the evidence is, the order of work, the catalogue of reds that are already understood, and the rules that keep an investigation from destroying its own evidence. Use when a build or a test goes red.
---

# headless-failure — finding out what went wrong

A red here is one of three things, and naming which one comes before any
reading of code.

| | What it looks like |
|---|---|
| **The run never started** | no output file, no report; the fault is in staging, import or compile |
| **The run finished and the output is wrong** | a file on disk with missing labels, no geography, a project that will not compile |
| **The run answered OK and recorded errors** | `[FAIL] the build answered OK and the generation recorded N error(s)` |

The third is the common one and it is usually the setup's own content.

## Where the evidence is

**A headless build**, under the output folder given to `--out`:

| File | What it answers |
|---|---|
| `<name>-build-report.txt` | `outcome=`, `sheets=`, `variables=` — the verdict and the size of what was built |
| `<name>-generation.txt` | the `[Error]` lines, with the sheet and the variable each belongs to |
| `<name>-import.log` | whether every component reached the project, and which one refused |
| `<name>-trace.txt` | the step the build was on |
| `<name>-designer.xlsm` | the designer copy the build actually drove |

**A batch**: `logs/run-<stem>.log` holds one run, `logs/batch-verdicts.txt`
holds one line per setup with the exit code, the file written, the sheet and
variable counts and the outcome. Read the verdict line before opening any log.

**A unit test run**: `.test-runner/tests/staging/test-results.csv`. Failed
assertions are typed `"error"`; the runner's own tally sits in the
`__summary__` row. A red run keeps its working copy and prints the path, and
that copy holds whatever partial output Excel managed even when the CSV never
printed.

**A verify case**: `.obt/draft/verify/run-<stamp>/<case>.log`. The `BAD` lines
name the assertions that did not hold, and a failed run keeps every file it
generated.

Read the report and the verdict line first. They answer "what happened" in one
line, and most investigations end there.

## The order of work

**1. Separate a product fault from a fixture fault.** Check the catalogue
below before anything else. Several reds in this repository are the correct
behaviour over deliberately broken input.

**2. Check intent before touching production code.** When the failure traces
to a method changed recently, run `git log -- <file>` and read the diff. The
owner iterates on production logic in small deliberate steps, several
revisions in a day. A test that fails after such a change is normally the test
lagging the new contract. Fix the test and leave the production code alone. The
owner has corrected exactly this, reverting a production "fix" with: "You
should fix the tests." Touch production code when the failure is a regression
the owner did not intend.

**3. Narrow by cutting the input.** The import of
five sheets answers one error number for the whole lot;
`convert-setup.ps1 -Sheets YYYYN` imports them cumulatively and names the sheet
that fails. The same idea works on a setup: build with the geobase removed, one
language, one sheet.

**4. Measure the input condition.** Write down the actual state the code met —
the range address, the row count, whether a `DataBodyRange` is `Nothing`. A
root cause is a condition plus the line that cannot handle it.

**5. Prove it by construction.** Reproduce the failure by making a clean input
carry that one condition. Deleting the Translations rows of a working setup
reproduced `ERROR 91` exactly, which is what turned an inference into a root
cause.

**6. Fix at the lowest place that owns the condition.** `BetterArray.cls` is a
no-go zone — fix at its caller. When a sibling procedure already guards the
same case, match the guard it uses, and keep the fix inline when it would serve
one place.

**7. Three attempts, then report and move on.** On a batch, a setup that has
failed three times gets written up with what was measured and the batch
continues. A long investigation of one setup blocks every other one behind it.

## The catalogue of understood reds

Recognising one of these saves an investigation.

| Signature | What it is |
|---|---|
| `[Error] ... brok_*`, "This formula should fail" | `demo_setup_filled` and `generic-test-setup` carry `DictionaryTestFixture.bas` fixtures on purpose. Their logs always hold `[Error]` lines and the launcher always exits 1. Correct behaviour on those two. |
| `malaria-aweil`: "The formula contains unmatched parentheses" on `case_id_auto` | the setup writes `IF(ISBLANK(case_id, "", …`, closing `ISBLANK` in the wrong place. A content fault in that setup. |
| `malaria-aweil`: `ERROR 91` on **Translations** | its `Tab_Translations` is `B4:E5` with `DataBodyRange Is Nothing`. Fixed in `CustomTable.ImportAll` on 2026-11-25. The conversion path drives `setup_dev.xlsm`'s own baked copy of that class, so the fix reaches it only after `Dev > Import Codes`. |
| `file access: one root, the call FAILED, error 438` | `GrantAccessToMultipleFiles` exists for the macOS sandbox and raises 438 on Windows. Harmless here. |
| `modulesRun=0` | `--build` was left off the test run. |
| the opaque `-50` | the compile closure is broken. A narrowed registry commented a `classes:` or `modules:` row. |
| VBE import `40040` | an LF-only `.cls` or `.bas`. Run `unix2dos` and relaunch. |
| error `1004` on `FreezePanes` or a `SheetView` write | a hidden Excel refuses these. Production logs it and carries on. |
| a green heading on a brand-new suite | a suite that fails to compile reports nothing and the runner still says green. Read the CSV rows and the import log. |
| `getaddrinfo() thread failed to start` | the git HTTPS flake. Retry two or three times — see `/commit-work`. |

**Old and current setups are told apart by their own named ranges.** Current
carries `RNG_tab_<table>_<field>_updated` and `__OBT__SILENT_OPERATIONS__`; old
carries `RNG_<field>_<table>` (for example `RNG_main_label_dict`) and
`RNG_DevPasswd`. Reading `xl/workbook.bin` out of the zip answers this in
milliseconds with no Excel.

**A setup import runs inside the target workbook.** `SetupImport` resolves its
destination as `ThisWorkbook` in seven places, so the injected
`OBTSetupImportHeadless.bas` executes in the copy being filled. A fix in `src/`
reaches that path only after the binary is refreshed.

## Rules that protect the evidence

- **Do not edit a file while a run that reads it is in flight.** Rscript reads
  and evaluates expression by expression, so an edit lands under the running
  interpreter and the run dies on whatever the offsets now point at. On
  2026-11-25 a comment edit to a header killed a build four minutes in with an
  error naming a line 600 lines away, and the file parsed clean afterwards.
  Builds take two to eight minutes, which is just long enough to tempt a
  session into useful work on the same files. Serialize: finish the edits, then
  launch. During a run, write only to the scratchpad and to the output folders
  no run reads.
- **Read the workbook's bytes before opening Excel.** `unzip -p`, `strings`,
  and `scripts/devtools/vba-inspect.R` answer most questions about a delivered
  file in a second, with no COM and no risk to a live run.
- **Judge the delivered file.** A build answers `OK` over a
  linelist with no labels and a project that will not compile. Run
  `Rscript scripts/devtools/linelist-inspect.R` on the output.
- **Never kill Excel by image name.** Quit the instance this session created,
  through its own COM object. A peer's run shares this machine.
- **No dialogs.** A modal dialog in an unattended run is a hang. Pass
  `Worksheet.Unprotect` an explicit string — a bare `Unprotect()` on a
  password-protected sheet prompts, and a string makes it raise instead, which
  a `catch` handles. The verify skill's "A case never puts up a dialog" section
  has the detail.
- **A rejected Excel run stays rejected.** When the owner declines one, finish
  from the evidence in hand and name what was left unmeasured.
- **Report what was measured and what was assumed**, separately. An inferred
  cause gets labelled as one.

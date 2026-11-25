---
name: verify
description: Run the behavioural verify suite before a release. Generates real linelists with real Excel and judges whether the output is coherent. Use when the owner says prepare a release, verify, run the edge cases, or before pushing binaries. Also read this before adding an edge case.
---

# verify — the behavioural suite

**Run this before a release.** It generates real files with real Excel and
judges what came out.

```bash
bash scripts/verify/verify.sh            # every case
bash scripts/verify/verify.sh --list     # name them, run nothing
bash scripts/verify/verify.sh --case 030 # the ones matching 030
bash scripts/verify/verify.sh --keep     # keep the generated workbooks
```

## What this answers, and what answers the rest

| | |
|---|---|
| `scripts/tests/run-tests.R` | the **unit** suites, inside Excel, over `src/` |
| `scripts/release/preflight.sh` | the **static** release gate: envelopes, ribbons, hand-imported drift, closure, translations |
| `scripts/verify/verify.sh` | **this**: the product's behaviour, end to end |

The preflight asks whether the binaries ship correctly. This asks whether the
thing they build is any good. A release wants both, and neither answers the
other's question.

A case here never tests a function in isolation. It builds a linelist and asks
whether a person opening it would find it coherent: does it read in a language,
does it carry its geography, does it hold the sheets its setup asked for, does
its project carry the classes that must travel.

## Before you run it

**Close Excel.** Every case launches its own Excel through the headless build,
and a watchdog stops a build that runs away on memory. The watchdog refuses to
start while another Excel is open, because it cannot tell a runaway build from
somebody's own workbook. The suite then exits 2 and says so.

`--no-guard` runs without the watchdog. On this laptop that is how a build once
hung the machine, so use it only when you have a reason and the owner knows.

The watchdog comes from `.obt/draft/headless/guard-start.ps1`, which starts the
same one `guard-ctl.ps1` starts and skips only its 10 GB free-commit
pre-flight. That check never passes here with the owner's own windows open.

**Read the watchdog lines the run prints at the end, every time.** As of
2026-11-25 the guard dies about a second into a run on this laptop and the rest
of the suite goes unwatched:

```
15:25:46.7  guard started cap=1500MB floorCommit=8GB floorRam=1.2GB deadline=7200s
15:25:47.6    commitFree=7.1GB ramFree=1.3GB
15:25:47.6  TRIPPED: free commit at 7.1GB, under the floor of 8GB
samples: 1
```

`guard-start.ps1` forwards `CapMB`, `DeadlineSeconds` and `FloorRamGB` and
leaves `FloorCommitGB` at `guard.ps1`'s 8 GB default. Free commit here starts
around 7 GB, so the first sample trips, `guard.ps1` calls `Stop-Owned` and
exits 3. The runner has already printed `watchdog up` and never learns it died,
so the cap, the RAM floor and the deadline are all gone while the suite reads
as guarded. `samples: 1` is the tell.

It killed nothing those times only because no Excel was registered yet. Had one
been, it would have been stopped a second in. Until `guard-start.ps1` passes
`-FloorCommitGB` down, treat a run as unwatched and do not leave the machine.

## Reading the result

One line per case, and the exit code is the verdict of the run.

```
010-generic-headless-build        PASS    63s  the generic setup builds and ...
030-empty-language-falls-back     FAIL   154s  1 assertion(s) did not hold
040-old-setup-import-lands-...    PASS   248s  a damaged setup lands what it can ...
```

| Verdict | What to do |
|---|---|
| **PASS** | nothing |
| **SKIP** | read the reason. A fixture is absent, or present but too stale to use, so that behaviour was **not tested**. Before a release, get the fixture and re-run, or tell the owner which coverage is missing. |
| **FAIL** | open `.obt/draft/verify/run-<stamp>/<case>.log`. The `BAD` lines name the assertions that did not hold, and a failed run keeps every file it generated. `/headless-failure` carries the investigation order and the catalogue of reds that are already understood. |

A green run deletes the workbooks it made and keeps the logs. A failed run
keeps everything, because what a broken build left behind is the only account
of what went wrong.

**SKIP is not a pass.** Every fixture is a binary workbook and `.gitignore`
keeps all of them off git, so a clone has the scripts and none of the inputs. A
suite that went red on a missing fixture would be red on every machine but one,
and people learn to ignore a red like that. Treat a SKIP before a release as
coverage you do not have.

## When the owner asks for a release

1. `bash scripts/verify/verify.sh` — this suite, the behaviour over `src/`.
2. `bash scripts/release/preflight.sh --check` — the static gate.
3. `bash scripts/verify-package/verify-package.sh` — the **package** over the
   released binaries, with data put through the result. Use `/verify-package`;
   it gates on the installed obt and the channel both being current, so a
   skip there means that question went unasked.
4. Report all three, and name every SKIP.

Never push binaries over a failing case or an unexplained SKIP without saying
so plainly. `RELEASING.md` carries the rest of the release.

## The fixtures

Cases resolve their inputs under `.obt/draft/headless/`, which
`OBT_VERIFY_FIXTURES` moves. `.obt/draft/headless/headless-progress.html`
records how that folder was built and what each setup is.

| What | Resolved from |
|---|---|
| the designer | `.mock/designer_mock.xlsm`, then `.mock/xlsb/designer_mock.xlsb` |
| the ribbon template | `ribbons/_ribbontemplate_dev.xlsm`, then the `.xlsb` |
| the empty setup | `src/bin/setup/setup_dev.xlsm`, then the fixtures' `template/` |
| the generic setup | `src/tests/.input/package/generic-test-setup.xlsb`, then `setups/old/generic.xlsb` |
| old and current setups | `setups/{current,old}/` under the fixtures |

Every fixture here is a workbook in the repository or the headless folder. **No
case in this suite needs an obt channel**, which is the point of the next
section.

### The package is judged by a different suite

**Nothing here drives the obt package or the released binaries. Use
`/verify-package` for that**, and do not add a case here that needs a channel.

There used to be one: case 020 built the generic setup through obt. It was
removed on 2026-11-25, because a case needing a release channel can only ever
judge whatever zip happens to be cached on the machine. A dev channel is stamped
`latest` whatever its age and obt accepts `latest` without looking further, so
such a case is either skipping on a stale fixture or quietly passing over
months-old binaries. Neither is coverage, and it made this suite's result table
read worse than the suite deserved.

`scripts/verify-package/` asks that question properly: it gates on the installed
obt matching its source and on the channel being one obt can drive, then drives
every current setup — generate, fake a dataset, import it back — and prints how
far the channel trails `src/` on every line. A release wants both suites.

## The cases

Read `--list` for the current set. Each case file opens with why it exists and
what it deliberately does not claim; read that before changing one.

## Adding a case

Use `/edge-case`. It asks what the behaviour is, writes the case against the
contract below, and keeps the suite's shape.

### The contract a case keeps

A case is `scripts/verify/cases/NNN-slug.sh`, executable, sourcing
`../lib/verify-lib.sh`.

```bash
#!/usr/bin/env bash
# title: one line, what the behaviour is
# asserts: one line, what is actually checked
```

The runner reads those two headers for `--list` and the result table.

| Exit | Meaning |
|---|---|
| 0 | **PASS**, through `check_case` or `pass_case` |
| 2 | **SKIP**, through `skip_case "<reason>"` |
| other | **FAIL**, through `check_case` or `fail_case` |

From the lib: `need_file`, `need_command` for preconditions, which **skip**
rather than fail; `first_existing` to resolve a fixture in preference order;
`assert_file`, `assert_contains`, `assert_not_contains`, `assert_equal`,
`assert_at_least`, `assert_file_bigger_than` which record and carry on;
`build_linelist` and `inspect_linelist` for the two wrappers around a build;
`strings_of` to read a delivered workbook's labels out of its zip with no
Excel; `report_value` to read a build report key; `excel_count` and
`wait_for_no_excel` to catch an operation that left Excel running;
`BROK_FIXTURE_ERRORS` for the deliberately broken formulas. `check_case` is the
last line.

### A case never puts up a dialog

**The suite runs unattended. A modal dialog is a hang somebody has to answer by
hand, and the owner should never be typing a password into a verify run.**

The trap that caught this suite once: `Worksheet.Unprotect()` **with no
argument** on a sheet that has a password makes Excel show the password dialog.
Passing a string — even the wrong one — makes Excel raise instead, which a
`catch` handles silently. Every one of a setup's five import sheets is
protected, and `Analysis` and `Translations` hold a password.

```powershell
$sh.Unprotect()           # WRONG: prompts when a password is set
$sh.Unprotect("$pw")      # right: raises on a wrong password
$sh.Unprotect("")         # right: the no-password case, still explicit
```

A workbook's own password is read off the copy being worked on, so nothing is
hardcoded and nothing is asked. It is the **workbook-level name
`RNG_DebuggingPassword`**, which is where the product keeps it —
`Passwords.cls` holds it as `NAMEDEBUGPASSWORD` and reads it through
`ThisWorkbook.Names`. Call `Get-DebugPassword <wb>` from `lib/excel-helpers.ps1`
rather than reaching for a cell.

**Follow the name, never a cell address.** The helper read `__pass!A1` until
2026-11-25. A1 happens to hold the password in the *old* setups, so case 060
lifted `Translations` on `ebola-2026-11-19.xlsb` and nothing looked wrong. On
the *current* setups A1 holds something else, both lift attempts failed, and
case 030 skipped with "the Translations sheet is protected and would not
unprotect" — which reads as a product fault and was a fixture bug. The name's
cell is not fixed, so an address is the wrong handle for it.

The same rule covers every other COM call that can prompt: pass
`Workbooks.Open` an explicit password argument rather than letting it ask, and
keep `DisplayAlerts = $false` and `EnableEvents = $false` on any Excel a case
creates.

Prefer not needing the lift at all. Read a workbook's bytes with `strings_of`
or `unzip` where the question allows it.

### Four rules a case follows

**Allow the expected errors by name, never in bulk.** `demo_setup_filled` and
`generic-test-setup` both carry `DictionaryTestFixture`'s `brok_*` variables,
labelled "This formula should fail". Their generation logs always hold
`[Error]` lines and `build-linelist.R` always exits 1 on them. Allow exactly
those through `BROK_FIXTURE_ERRORS`. An allowance of "any error" passes a build
that broke everything else.

**Judge the output, not the exit code.** A build answers `OK` over a linelist
with no labels, no geography and a project that will not compile. Assert on the
delivered file: `inspect_linelist` for the faults the repository already knows
how to name, and the workbook's own bytes for anything else.

**Say what the case does not claim.** Case 040 writes the gap it leaves in its
own header. A case that quietly asserts less than the requirement is worse than
no case, because it reads as coverage.

**Never pin a bug, and manufacture the fault rather than finding one.** Case 040
converted `malaria-aweil.xlsb` because that setup happened to raise error 91 on
its zero-body `Tab_Translations`, and asserted that the import reported a fault.
Case 060 then covered that bug and it was fixed — so 040 went red for the best
possible reason, and the coverage it was carrying died with it. Two lessons, and
they apply to every case here:

- **Build the damaged input, don't wait for one.** `cases/040-delete-rows.ps1`
  deletes a table's rows, or the whole table, on a throwaway copy. A fixture cut
  from a setup that is broken by accident expires the day somebody fixes it.
- **Assert the disjunction when the product might improve.** 040 requires the
  import to answer `OK` **or** a reported error, lands-what-it-can either way,
  and then *prints* which branch happened. Asserting the silence would pin
  today's bug and go red on the fix. Asserting the report would go red today.
  The run says which, and the header says which is live.

Right now that branch is the silent one: a source whose `Tab_Translations` has
been deleted outright converts with outcome `OK`, 381 dictionary rows, and no
mention of the missing table. "What cannot land is reported" is a requirement
the product does not yet meet, so no case asserts it.

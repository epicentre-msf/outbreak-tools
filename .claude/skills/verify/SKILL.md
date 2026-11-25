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
pre-flight. That check never passes here with the owner's own windows open. The
memory cap, the RAM floor and the deadline all stay.

## Reading the result

One line per case, and the exit code is the verdict of the run.

```
010-generic-headless-build        PASS   184s  the generic setup builds and ...
020-generic-obt-package           SKIP    0s   no channel zip cached on this machine
030-empty-language-falls-back     FAIL  298s  1 assertion(s) did not hold
```

| Verdict | What to do |
|---|---|
| **PASS** | nothing |
| **SKIP** | read the reason. A fixture is absent, so that behaviour was **not tested**. Before a release, get the fixture and re-run, or tell the owner which coverage is missing. |
| **FAIL** | open `.obt/draft/verify/run-<stamp>/<case>.log`. The `BAD` lines name the assertions that did not hold, and a failed run keeps every file it generated. |

A green run deletes the workbooks it made and keeps the logs. A failed run
keeps everything, because what a broken build left behind is the only account
of what went wrong.

**SKIP is not a pass.** Every fixture is a binary workbook and `.gitignore`
keeps all of them off git, so a clone has the scripts and none of the inputs. A
suite that went red on a missing fixture would be red on every machine but one,
and people learn to ignore a red like that. Treat a SKIP before a release as
coverage you do not have.

## When the owner asks for a release

1. `bash scripts/verify/verify.sh` — this suite.
2. `bash scripts/release/preflight.sh --check` — the static gate.
3. Report both, and name every SKIP.

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

A workbook's own password is read off the copy being worked on —
`__pass!A1` — so nothing is hardcoded and nothing is asked. The same rule
covers every other COM call that can prompt: pass `Workbooks.Open` an explicit
password argument rather than letting it ask, and keep `DisplayAlerts = $false`
and `EnableEvents = $false` on any Excel a case creates.

Prefer not needing the lift at all. Read a workbook's bytes with `strings_of`
or `unzip` where the question allows it.

### Three rules a case follows

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

**Say what the case does not claim.** Case 040 asserts today's behaviour and
writes the gap it leaves in its own header. A case that quietly asserts less
than the requirement is worse than no case, because it reads as coverage.

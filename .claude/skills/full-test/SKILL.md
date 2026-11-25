---
name: full-test
description: Run the full VBA unit test suite, or the narrowed probe that comes before it. Covers the scan gate, the registry narrowing, the memory watchdog, how to launch a run that takes two hours, and how to read the result. Use when asked to run the tests, probe a change, or check a suite before a release.
---

# full-test — the unit suite inside Excel

One command drives the whole suite over `src/tests/test-registry.yml`:

```bash
Rscript scripts/tests/run-tests.R --build
```

Pass `--build` unless the workbook's Codes tables are known current. Without it
the run imports nothing and reports `modulesRun=0`.

| | |
|---|---|
| `scripts/tests/run-tests.R` | **this**: the unit suites, inside Excel, over `src/` |
| `/verify` | the product's behaviour end to end, over generated files |
| `scripts/release/preflight.sh` | the static release gate |

## A full run is the owner's call

**Do not launch a full run unprompted.** A session closes on its narrowed
probe, writes "full run owed" in `.obt/tracking.md`, and stops there. This
holds even when `implementations.md` or `conventions.md` says a full run is
wanted for the change in hand.

The reason is cost. The last green full run (2026-11-03) ran 09:58 to 11:50,
about **1h50m** over 114 modules, 2030 tests and 6123 assertions. It holds the
one hidden Excel this machine can give for that whole time, so every other
probe is blocked behind it. A narrowed probe answers the same question for the
touched folders in two to four minutes.

A full run is wanted for: a registry change that adds or removes a module, a
change to a class many folders name (`Checking`, `BetterArray`, `HiddenNames`,
`CustomTable`, `LLFormat`), a change to the harness itself, and before a
release. Those are the cases where the untouched modules stopped being
untouched.

## 1. The scan gate, before any run

All three must pass. They catch the compile faults a run would otherwise die
on forty minutes in.

```bash
Rscript scripts/devtools/call-signature-scan.R
Rscript scripts/devtools/name-shadowing-scan.R
bash    scripts/devtools/interface-guard.sh
```

Then `unix2dos` every `.cls`, `.bas` and `.yml` a shell edit touched. The
Windows runner's `OBTSilentImport` refuses an LF-only component with error
40040, and a run launched over one dies in the import step.

## 2. The narrowed probe

Let the script make the cut. Commenting a `- module:` row by hand and leaving
its `covers:` continuation standing kills the YAML on a duplicate key.

```bash
Rscript scripts/devtools/narrow-registry.R <folder>[,<folder>...]
Rscript scripts/tests/run-tests.R --build
git checkout -- src/tests/test-registry.yml
```

Name several folders at once to reproduce a failure that needs both of them: a
suite that leaves state behind only reaches its victim when both are in the
same run.

Only `- module:` rows under a `tests:` key are commented. Every `classes:`,
`modules:`, `fixtures:` and `forms:` row stays, because the compile closure has
to stay whole. A broken closure makes the project refuse to build and the run
answers the opaque `-50`.

**Never commit a narrowed registry.** Restore it with `git checkout --` so it
comes back byte-identical for any other session reading it.

**Report what the probe measured.** Give the module count and the assertion
count, and name the suites that were in it. A probe count is a statement about
those suites alone.

### When a peer session shares the checkout

The owner runs several sessions at once. Probe under a private home, with the
driver workbook copied in and `tests/staging` made:

```bash
Rscript scripts/tests/run-tests.R --build --home=.test-runner/<session>-home
```

Copy `src/tests/test-registry.yml` aside before narrowing and back after, md5
checked. A peer's `git checkout -- <registry>` takes whatever is in the index,
so stage the rows early.

## 3. The watchdog, before the Excel starts

Every run on this laptop goes up under a detached memory watchdog. An unbounded
probe once hung the machine hard enough to need a restart.

```powershell
.obt\draft\headless\guard-start.ps1 -Tag <tag> -DeadlineSeconds 7200
.obt\draft\hidden-instance-probe\guard-ctl.ps1 -Action show -Tag <tag>
.obt\draft\hidden-instance-probe\guard-ctl.ps1 -Action stop -Tag <tag>
```

`guard-ctl.ps1 -Action start` refuses to begin under 10 GB of free commit,
which never passes with the owner's own windows open. `guard-start.ps1` skips
that one gate and keeps every live protection: the cap on Excel's private
memory, the free-RAM floor and the deadline.

Both refuse to start while any Excel is already running, because they cannot
tell a runaway build from somebody's own workbook. Close Excel first.

**Read the guard log every time.** `samples: 1` means it tripped on its first
sample and died a second in, and everything after that ran unwatched:

```
15:25:47.6  TRIPPED: free commit at 7.1GB, under the floor of 8GB
samples: 1
```

`guard-start.ps1` leaves `FloorCommitGB` at `guard.ps1`'s 8 GB default and free
commit here starts near 7 GB, so that first sample trips. Until
`guard-start.ps1` forwards `-FloorCommitGB`, treat such a run as unwatched and
stay with the machine.

### Four rules the watchdog work earned

- **The deadline is wall-clock.** It keeps counting through system sleep. On
  2026-10-29 a run's idle Excel was correctly killed on wake, four real days
  after launch. Confirm the laptop will stay awake (the owner uses Caffeine)
  before launching anything over about 30 minutes unattended.
- **Set a generous deadline from the start.** 7200s covers a full run with
  room. Extending a live watchdog mid-run is where the accident below happened.
- **Never point a debug invocation at a live watch file, and never pair a
  `-TrustedPids` override with a short deadline.** On 2026-10-02 a throwaway
  `-DeadlineSeconds 5` syntax check reused the real watch file; the trusted PID
  bypassed the "started after this watchdog" ownership check and the 5-second
  deadline killed a real test run 35 minutes in. Use a fake PID or a scratch
  watch file for syntax checks.
- **Prove a new watchdog on an idle Excel first**, with a cap any Excel is
  over, before pointing it at a run.

A hidden Excel always gets `ScreenUpdating = False`. With it on, an invisible
instance holds about 0.7 MB of shared GPU memory per cell write outside its
private bytes — 5,000 writes took 3 GB in four seconds — so private bytes stay
flat while the machine runs out of RAM.

Excel is stopped **by PID through the owned COM object**. Killing by image name
takes down a peer's run on the same machine; the memory rule
`never-kill-excel-by-name` governs this.

## 4. Launching a run that lasts hours

The run writes thousands of inventory lines and takes longer than any
foreground tool call, so it goes to disk, detached. The shape used on Windows:

```
.test-runner/run-full-<stamp>.cmd      cd /d <repo> && Rscript run-tests.R --build > <log> 2>&1
.test-runner/full-run-<stamp>.log      the whole run
```

Start the `.cmd` detached through WMI (`Invoke-CimMethod Win32_Process
Create`), the way the watchdog starts. Started as a child of the tool call, an
interruption ends the driver and leaves the COM-launched Excel alive and
hidden.

On Windows the test Excel runs hidden (`scripts/tests/windows/run-tests.vbs`,
`Visible = False`). A hidden Excel refuses `Window.FreezePanes` and `SheetView`
writes with error 1004; production logs the refusal and carries on, and no test
asserts a frozen pane.

`scripts/tests/run-tests-bg.sh` is the macOS wrapper for the same idea.

### Before relaunching

Kill the leftover `EXCEL`, `cscript` and `Rscript` processes **this session**
started, by PID, and delete `.test-runner/tests/staging/run`. A stale copy that
survived an aborted run crashes the next one, and an operator VBE sitting in
break mode holds that copy locked.

**Do not edit a file while a run that reads it is in flight.** Rscript reads
and evaluates expression by expression, so an edit lands under the running
interpreter. On 2026-11-25 a comment edit killed a build four minutes in with
an error naming a line 600 lines away. Finish the edits, then launch.

## 5. Reading the result

```
run-tests.R: summary — total=6123 failures=0 | modulesRun=114 testsFound=2030 vbe=ok(248)
run-tests.R: 6123 success / 0 failure row(s) across 6124 result line(s).
run-tests.R: all tests passed.
```

Exit 0 on all-pass, 1 on any failure or an incomplete run. A green run deletes
the per-run working copy unless `--keep`; a red one keeps it and prints the
path, and that copy holds whatever partial output Excel managed even when the
CSV never printed.

| Reading | What it means |
|---|---|
| `modulesRun=0` | `--build` was left off; nothing was imported |
| exit 1, no summary line | read the import log in the kept run dir |
| the opaque `-50` | the compile closure is broken; a registry row was commented that should not have been |
| a green heading on a new suite | **check it.** A suite that fails to compile reports nothing and the runner still says green |

Results land in `.test-runner/tests/staging/test-results.csv`. The runner writes
its own tally into a `__summary__` row; read that row for the count, because a
hand count and the runner's number can only ever disagree. A failed assertion
is typed `"error"` in that file.

**Read the CSV rows and the import log before trusting a green heading on a
suite that has just been added.** That is the one green that lies.

When a failure traces to a method changed earlier the same day, go to
`/headless-failure` before touching production code. The owner iterates on
production logic in deliberate steps and the test is usually the thing lagging.

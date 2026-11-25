---
name: headless-generation
description: Run the whole headless pass — pull the working folder, convert the old setups into the current structure, generate a linelist for every setup, write the results into headless-progress.html with the date and the OS, and report what the reds were. Use when the owner says run the headless generation, do the headless pass, regenerate the linelists, check the old setups still convert, or asks whether every setup still builds.
---

# headless-generation — the whole pass, one command

Five questions, in this order, and each one depends on the last:

| | |
|---|---|
| 0 | is this machine's working folder the current one |
| 1 | do the **old** setups still convert into the current structure |
| 2 | does every setup, current and converted, still **generate** a linelist |
| 3 | what does `headless-progress.html` say about this run |
| 4 | what should the owner know about doing this on the other OS |
| 5 | which reds are real, and what was done about them |

Nothing here commits, rebuilds a binary or runs a unit suite. `/full-test` and
`/commit-work` own those, and both belong after this, not inside it.

## What it costs, before you start

A full pass on Windows is **about 25 minutes of Excel**: six current setups in
roughly 11 minutes, five converted ones in 7, the conversions themselves 14 to
23 seconds each. On macOS the same work has run anywhere from 185 to 749
seconds **per build**, so budget an hour and read the macOS section first.

Say that number to the owner before launching, and say which setups are in the
pass. A pass they did not expect to wait 25 minutes for is a pass they will
interrupt halfway.

## Before anything launches Excel

- **Close Excel.** The watchdog refuses to start while one is running, because
  it cannot tell a runaway build from somebody's own workbook.
- **One watchdog covers the whole pass**, Windows only, started before the
  first Excel and stopped at the end:

  ```powershell
  .obt\draft\headless\guard-start.ps1 -Tag headless-pass -DeadlineSeconds 7200 -FloorCommitGB 4
  .obt\draft\hidden-instance-probe\guard-ctl.ps1 -Action show -Tag headless-pass
  .obt\draft\hidden-instance-probe\guard-ctl.ps1 -Action stop -Tag headless-pass
  ```

  **Lower the commit floor on this laptop, or the guard stops a healthy build.**
  The floor defaults to `guard.ps1`'s 8 GB and free commit here starts at or
  below that, so the guard trips within seconds on `free commit at 8.0GB, under
  the floor of 8GB`. `-FloorCommitGB` reaches it as of 2026-11-25; before that
  nothing forwarded the bound. Every live protection stays: the cap on Excel's
  private memory, the free-RAM floor and the deadline. Lowering the default is
  the owner's call, so pass it per run.

  **Read its log either way**, at
  `.obt/draft/hidden-instance-probe/run-<tag>-guard.log`. `samples: 1` means it
  tripped on its first look and died a second in, and everything after ran
  unwatched — treat such a pass as unwatched, stay with the machine, and say so
  in the report. **A guard trip can arrive wearing the wrong label**: a killed
  Excel surfaces as `osascript trigger returned 1 (Excel may have wedged on a
  dialog)`, which points at a dialog that was never there. On that message, read
  the guard log before anything else, and a low sample count is the tell.
- **Finish every edit first.** Rscript evaluates as it reads, so an edit to
  `build-linelist.R` or anything under `src/` lands under the running
  interpreter and kills the run on an offset 600 lines from the change. While a
  pass is in flight, the only safe places to write are the scratchpad,
  `out/`, `logs/` and `headless-progress.html`.
- **Excel is visible for the length of a build, by design.** A hidden Excel
  refuses `Window.FreezePanes` with error 1004 and the build freezes the header
  of every data entry sheet. Leave it alone.
- **Never kill Excel by image name.** A peer session shares this machine. Quit
  the instance this pass created, by PID, through its own COM object.
- **One Excel at a time.** This laptop is memory-tight; `run-batch.sh` already
  waits for one to go before launching the next.

## Step 0 — the working folder

The inputs are not in git. `.obt/draft/headless/` is untracked and holds the
setups, the geobases, the empty setup template and the drivers, and it moves
between machines through a private store of its own.

```bash
bash scripts/headless/pull-headless.sh --check
```

| Answer | What to do |
|---|---|
| `IN SYNC` | nothing. Go to step 1 |
| `AHEAD` | `bash scripts/headless/pull-headless.sh`, then say in the report that the folder came from the store and name the stamp's date and OS |
| `EMPTY` | nothing has ever been pushed. On a machine that **has** the folder, push it (`bash scripts/headless/push-headless.sh`) and carry on. On one that does not, stop: there is nothing to generate from, and say exactly that |

A pull overwrites the inputs and leaves `out/` and `logs/` alone, because the
bundle does not carry them. `scripts/headless/README.md` has the rest.

Then confirm what the pass will walk, which launches nothing:

```bash
bash .obt/draft/headless/run-batch.sh --all-current   --list
bash .obt/draft/headless/run-batch.sh --all-converted --list
```

## Step 1 — the old setups convert

**Windows only.** `convert-setup.ps1` drives Excel through COM and PowerShell
and has no macOS twin. On a Mac, skip this step, say it was skipped, and
generate from the `setups/converted/` the store already carries.

```powershell
.obt\draft\headless\convert-all.ps1
.obt\draft\headless\convert-all.ps1 -Only ebola-2026-11-19,generic
```

Each conversion fills a **fresh copy of the current empty setup** from the old
one; the old file is opened read-only and never changed. Verdicts land one line
each in `logs/convert-verdicts.txt`, so read that before any log.

A conversion is green when it answers `OK` **and** the dictionary row count is
above zero. The import answers `OK` over an empty dictionary as readily as over
a full one, which is why the row count is part of the verdict and not a detail.

`malaria-aweil` is the standing partial: `ERROR 91` on its **Translations**
sheet, four other sheets landing. Convert it with `-Sheets YYYYN` and move on —
`/headless-failure` carries the measurement.

## Step 2 — every setup generates

```bash
bash .obt/draft/headless/run-batch.sh --all-current
bash .obt/draft/headless/run-batch.sh --all-converted
```

Both enumerate the folder, so a setup added by `/add-setup` rides along with no
edit here, and `setups/geo-map.csv` decides which geobase each stem gets.

**Launch each batch in the background and wait for it to exit.** A current-setup
batch runs past the foreground tool limit, and polling a log while Excel builds
wastes the machine. Launch with `run_in_background`, write stdout to
`logs/batch-<stamp>.out`, and wait for the notification. Do not start the second
batch until the first has exited — two Excels is how this laptop falls over.

### What green means

**A generation is green when the build passes and delivers the file.** Judge it
on three things:

- `outcome=OK` in `out/<stem>/<name>-build-report.txt`
- the linelist on disk
- `sheets=` and `variables=` counts that are not zero

**Reported formula errors are not part of the verdict.** The `[Error]` lines in
`<name>-generation.txt` are the product telling a setup's author what is wrong
with their setup. Report them, count them, never colour them. The owner set
this on 2026-11-25: `demo` and `generic` both carry the `brok_*` variables from
`DictionaryTestFixture.bas`, labelled "This formula should fail", so a rule that
reds on reported errors makes those two permanently red and teaches everyone to
ignore the colour.

**The exit code carries the older rule.** `build-linelist.R` exits 1 whenever
the generation recorded an error, so `exit=1` with `outcome=OK` and a delivered
file is **green**. `run-batch.sh` judges the report and the file and prints the
exit code beside them, so `GREEN exit=1 errors=2` is the normal shape for a
setup carrying deliberate formula errors, and `malaria-aweil` is the standing
example.

**This file is where the rule lives.** A change to it belongs here first;
`/headless-failure` restates it compactly and points back. Settled with the peer
session on 2026-11-25.

**`logs/batch-verdicts.txt` holds two rules under one heading.** The blocks
written before 2026-11-25 red `demo`, `generic` and `malaria-aweil` on their
exit codes. That record was left as it was rather than rewritten, so **read a
verdict line against the date of its block** — and when quoting an old line in a
report, say which rule produced it.

## Step 3 — the page

`.obt/draft/headless/headless-progress.html` is the record both machines read.
Two regions of it are machine-written and **nothing else is**:

```
RUNSTAMP:BEGIN .. RUNSTAMP:END     the line under the title
RESULTS:BEGIN  .. RESULTS:END      the whole results section
```

Write both as fragments and splice them in. Base R, no dependencies, runs
wherever the build runs:

```bash
R_PROFILE_USER=nonexistent Rscript .obt/draft/headless/write-results.R \
  --results=<scratchpad>/results.html --stamp=<scratchpad>/stamp.html
```

**The results region is replaced wholesale, every run.** Results do not
accumulate: the page says what the last pass found. The run stamp is what says
*when* and *where*, so it carries the date, the OS and the Excel version, and
the counts behind the verdict:

```html
<p class="sub">
  <code>.obt/draft/headless/</code> &middot; branch <code>dev</code>
  &middot; last run <strong>2026-11-25</strong>
  &middot; <strong>Windows 11</strong>, Excel 16.0.20527
  &middot; 11 of 11 green
</p>
```

Read the Excel version off the machine rather than off the last run. On
Windows, from the Click-to-Run configuration, which is where this laptop's
`16.0.20527.20112` comes from:

```powershell
(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').VersionToReport
```

On macOS, `defaults read /Applications/Microsoft\ Excel.app/Contents/Info.plist CFBundleShortVersionString`,
or the About pane. A page stamped with the previous run's Excel is how a version
regression goes unnoticed, and 16.112 on macOS is exactly that failure.

The results region carries, in this order: one sentence of verdict, the
current-setups table, the converted-setups table, a table of the reported
errors naming who owns each, and what the conversions did. Keep the
`Errors` column and the `Sec` column — the first is how reported errors stay
visible without colouring a run, the second is how the next session knows
whether a build got slower.

Anything worth keeping across runs goes **outside** the markers: a new gotcha
into `Gotchas worth keeping`, a new macOS fault into `Generating on macOS`, an
open decision into `To fix and to improve`. The writer script refuses a page
whose markers are missing rather than guessing where the results went.

## Step 4 — the macOS caveats

The page's `Generating on macOS` section is the long version and should be read
before a pass on a Mac. Say these five out loud in the report, because they are
the ones that waste an afternoon:

1. **Excel 16.112 crashes every build at the first chart** — 5 runs out of 5,
   measured 2026-08-17 by downgrading Excel and changing nothing else. 16.111
   and 16.91 are both green. One release back, not twenty-one.
2. **A one-in-sixteen `ERROR 51`** at the first data entry sheet.
   `build-linelist.R` retries the trigger once on that signature, which takes it
   to roughly 1-in-250, and a retried run prints
   `NOTE, this run needed a second attempt` twice. Never read a retried run as a
   clean one.
3. **The sandbox decides where the run stages itself.** `OBT_HOME` defaults
   inside Excel's own container, where nothing is granted and nothing is picked,
   and the price is that **R needs Full Disk Access** to reach it. The
   `headless-runner` fallback needs a folder pick that covers exactly one build
   and is discarded at quit.
4. **No conversion and no watchdog.** `convert-setup.ps1` and `guard-start.ps1`
   are both PowerShell. Convert on Windows, push, and let the Mac pull
   `setups/converted/`; a macOS pass runs unwatched, so stay with it.
5. **Extension and content must agree** or Excel for Mac refuses the file
   outright, where Windows sniffs the content and opens it anyway. This is why
   both designer copies take the extension of the designer handed in.

`.obt/gotchas/headless-r-environment.md` and
`.obt/gotchas/macos-sandbox-grant.md` carry the evidence for all five.

## Step 5 — the reds, and three attempts each

Apply the green rule first. Most "reds" are a non-zero exit over a delivered
file, and those are green.

A build that died in its first seconds is a **watchdog trip** before it is
anything else, whatever message it carried out. Read
`run-<tag>-guard.log` first, then relaunch with a lower `-FloorCommitGB`.

For a real red — no report, no file, or counts of zero — go to
`/headless-failure` and work its order: name which of the two failures it is,
check the catalogue of understood reds, narrow by cutting the input, measure the
input condition, then fix at the lowest place that owns it.

**Three attempts, then write it up and move on.** An attempt is an
investigation plus a fix plus one re-run of that setup alone:

```bash
bash .obt/draft/headless/run-batch.sh <stem>              # current
bash .obt/draft/headless/run-batch.sh --converted <stem>
```

A fourth attempt on one setup blocks every other one behind it. After the third,
the setup gets a paragraph in the report and a row in the page's `To fix` list
saying what was measured, what was tried, and what the next session should look
at. Report it as unfixed, plainly. Do not report a workaround as a fix, and do
not quietly drop a setup from the pass.

### Which copy of the class actually ran

When a pass is judging a fix, this is the question to answer before reporting
anything, because the three paths execute three different copies of the same
class and they disagree silently.

| Path | Whose code runs | A `src/` fix is live |
|---|---|---|
| **a headless generation** | `src/`, re-imported into the driver every run | **immediately**, no binary refresh |
| **a conversion** | the setup workbook's own baked VBProject — `SetupImport` resolves `ThisWorkbook` in seven places, so the injected module calls what is already in there | only after `Dev > Import Codes` and a rebuild |
| **the unit suite** | the driver workbook's baked copy, for any class the registry forbids registering | **never**, until the workbook is rebuilt |

That third row is not hypothetical. `src/tests/test-registry.yml:294` says in as
many words not to add `TranslationObject` as a class, because the driver already
holds it and a second import leaves Excel keeping both as `TranslationObject1`
— so `.test-runner/unit_tests_dev.xlsb` carries a stale copy and no unit test
can see a `src/` change to that class. Measured 2026-11-25, alongside
`CustomTable` baked into `setup_dev.xlsm` on the conversion path. A probe over
such a class proves the test detects the bug and says nothing about the fix.

Two measurements answer it, and neither needs Excel:

```bash
# did this run import the class from src, or inherit a baked one?
grep -n '<ClassName>' .obt/draft/headless/out/<stem>/ll-<stem>-import.log

# how far a workbook's baked copies have drifted from src/
Rscript scripts/devtools/vba-inspect.R <path/to/vbaProject.bin> content src
```

`comp: TranslationObject | type=2 | lines=779` in an import log is the headless
build taking the class from source. Absence from that log means the run used
whatever the copied workbook had baked in.

So a report says **which copy ran**, and a fix verified on one path is reported
as verified on that path alone.

## When the pass is done

1. **Push the folder back**: `bash scripts/headless/push-headless.sh`. The page
   with this run's results is in the bundle; `out/` and `logs/` are not, so the
   other machine gets the record without the gigabytes.
2. **Stop the watchdog** and read its log one last time.
3. **Report**: one line per setup, the verdict, the counts, the reported-error
   counts with whose content they are, what the page now says, and anything left
   unfixed after three attempts. When the pass was run to judge a `src/` change,
   say which copy of the code each path executed — a generation reads `src/`,
   a conversion and the unit suite do not.

Say plainly what was not run. A pass that skipped the conversions because it ran
on a Mac is a pass that answered four of the five questions, and the report says
so.

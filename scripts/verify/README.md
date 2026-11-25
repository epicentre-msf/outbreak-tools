# The verify suite

The behavioural half of the release gate. It generates real linelists with real
Excel and judges whether what came out is coherent.

```
bash scripts/verify/verify.sh            every case
bash scripts/verify/verify.sh --list     name them, run nothing
bash scripts/verify/verify.sh --case 030 the ones matching 030
bash scripts/verify/verify.sh --help     every option
```

**Close Excel first.** Every case launches its own through the headless build,
and the watchdog that stops a runaway build will not start while another Excel
is open.

## What each thing answers

| | |
|---|---|
| `scripts/tests/run-tests.R` | the **unit** suites, inside Excel, over `src/` |
| `scripts/release/preflight.sh` | the **static** gate: envelopes, ribbons, import drift, closure, translations |
| `scripts/verify/verify.sh` | **this**: the product's behaviour, end to end |

The preflight asks whether the binaries ship correctly. This asks whether the
thing they build is any good.

## The layout

```
verify.sh                 the runner
lib/verify-lib.sh         what a case is given: paths, assertions, build wrappers
lib/excel-helpers.ps1     what a fixture builder needs from Excel, prompting never
cases/NNN-slug.sh         one behaviour each
cases/NNN-*.ps1 | .R      the fixture builders and drivers a case needs
```

**Nothing in here may put up a dialog.** The suite runs unattended, so a modal
is a hang somebody has to answer by hand. `Worksheet.Unprotect()` with no
argument prompts when the sheet has a password; `Lift-SheetProtection` in
`lib/excel-helpers.ps1` does it without ever asking, and every one of a setup's
five import sheets is protected.

Output goes to `.obt/draft/verify/run-<stamp>/`: one log per case, `results.txt`
and, on a failed run, every file that was generated.

## Exit codes

```
0  every case passed or skipped
1  at least one case failed
2  the suite could not run (usually: Excel is open, so the watchdog refused)
```

A **SKIP** is not a pass. Every fixture is a binary workbook that `.gitignore`
keeps off git, so a clone has the scripts and none of the inputs, and a case
with no fixture says so rather than going red on every machine but one. Before
a release, read every SKIP and decide whether you can ship without that
coverage.

## Adding a case

`/edge-case` in Claude Code, or write one by hand against the contract in
`.claude/skills/verify/SKILL.md`. A case declares two headers the runner reads:

```bash
# title: one line, what the behaviour is
# asserts: one line, what is actually checked
```

and ends through `check_case`. Preconditions skip; assertions record and carry
on, so one case reports every fault it found.

`.obt/draft/headless/headless-progress.html` records how the fixture folder was
built and what each setup in it is.

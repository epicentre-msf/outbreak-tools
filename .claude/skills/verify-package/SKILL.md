---
name: verify-package
description: Drive every headless setup through the obt package — generate its linelist, make a fake dataset from the same setup, import that dataset into the linelist. Runs only when the installed obt and the cached channel are both current. Use when the owner says verify the package, test the package, check obt end to end, or before shipping a release of the package.
---

# verify-package — the package suite

**This judges what a user installs.** The obt package, driving a downloaded
designer, with data put through the linelist that comes out.

```bash
bash scripts/verify-package/verify-package.sh              # every setup
bash scripts/verify-package/verify-package.sh --list       # name them, run nothing
bash scripts/verify-package/verify-package.sh --setup ebola
bash scripts/verify-package/verify-package.sh --records 50 # rows per worksheet
bash scripts/verify-package/verify-package.sh --keep       # keep what it generated
```

## What this answers, and what answers the rest

| | |
|---|---|
| `scripts/tests/run-tests.R` | the **unit** suites, inside Excel, over `src/` |
| `scripts/release/preflight.sh` | the **static** release gate |
| `scripts/verify/verify.sh` | the product's **behaviour**, end to end, over `src/` |
| `scripts/verify-package/` | **this**: the shipped **package** over the shipped **binaries** |

`verify.sh` builds with the mock designer carrying `src/` imported over it, so
it judges the code in the working tree. This one judges the release: obt fetches
a channel, drives `clickGenerate` on that designer, and the linelist is then
filled with records. Both matter and neither answers the other's question.

For each setup it does three things, and a failure names which:

1. `obt_designer_generate` — a linelist out of the setup
2. `obt_fake` — a dataset drawn from that same setup's dictionary
3. `obt_linelist_import` — that dataset into the linelist just generated

Then `linelist-inspect.R` judges the delivered file, the same judge `verify.sh`
uses, so "coherent" means one thing across both suites.

## The two gates

**Both must pass or nothing runs.** Each prints its reason and the command that
fixes it, then exits 2.

### 1. The installed obt must be the latest

A run over a stale obt reports on a package its author has already moved past,
and a red then reads as a product fault when the fix is written and merely not
installed here.

`gates/package-version.R` compares the installed package with its source
checkout — `OBT_SOURCE`, else the sibling `../obt`:

| Checked | Why both halves |
|---|---|
| `DESCRIPTION` Version | the readable half |
| installed `RemoteSha` vs source `HEAD` | the precise half: several commits land per version bump |
| source tree clean | an install came from a commit, so a dirty tree can never match one |

No source checkout makes the answer **unknown**, which also skips. Guessing
"yes" is how a stale install goes unnoticed.

Seen 2026-11-25: installed 0.0.1 at `c2687b02`, source at 0.0.2, two commits
ahead. Fix it with the `PKG_REINSTALL` line the gate prints.

### 2. The channel must be one obt can drive

obt reads the version off the designer's filename, everything after the first
`-`, and drives exactly two shapes:

| Filename | Version | obt |
|---|---|---|
| `designer_dev-latest.xlsb` | `latest` | drives it, unconditionally |
| `designer_main-2.0.0.xlsb` | `2.0.0` | drives it when `>= OBT_MINIMUM_VERSION` |
| `designer_dev-2026-06-11.xlsb` | `2026-06-11` | **aborts**: predates the semantic scheme |

Those older releases carry no entry point a script can call, so a run would stop
on a message box. `gates/channel-age.R` applies obt's own constants through
`obt:::` rather than copying the rule.

Build a current one with `bash scripts/release/build-release-zip.sh dev latest`,
or point `OBT_VERIFY_CHANNEL` at a folder holding a channel zip.

**The channel's age is printed, never a gate.** A dev channel is stamped
`latest` whatever its age and obt accepts `latest` without looking further, so
the version carries no age by design. The run compares the dates inside the zip
with the last commit touching `src/` and puts the lag on every result line:

```
ebola   PASS  142s  generated, faked and imported (channel 167d behind src/)
```

Read that as a pass for the binaries in that zip. Before shipping, rebuild the
channel from the branch being released.

## Why the import is forced

`obt_linelist_import` reads a **migration** file and warns on three things: no
metadata, no language recorded, or a different language. A person clicking the
button reads the warning and decides; a scripted run has nobody to decide, so it
stops and says which of the three stopped it.

`obt_fake` writes a plain workbook through `writexl::write_xlsx` — tables drawn
from the setup, with no migration metadata or language anywhere in it. So it
trips the first two warnings **by construction, every time**. `force = TRUE` is
structural here rather than a convenience, and the driver says so out loud.
`OBT_PACKAGE_FORCE=0` turns it off for anyone who wants to see the refusal.

## Before you run it

**Close Excel**, same as `verify.sh`: each setup launches its own through obt,
and the watchdog will not start while another is open.

**Then read the watchdog lines at the end.** `samples: 1` means the guard
tripped on its first look and exited, and everything after ran unwatched. On
this laptop it does that every run — `guard-start.ps1` leaves `FloorCommitGB` at
8 GB and free commit here starts lower. The run still prints `watchdog up`,
because the runner never learns the guard died.

## Reading the result

One line per setup, and the exit code is the verdict of the run.

| Verdict | What to do |
|---|---|
| **PASS** | nothing, but read the channel lag on the line |
| **FAIL** | open `.obt/draft/verify-package/run-<stamp>/<setup>.log`. The driver prints `PKG_STAGE_FAILED` naming the verb that raised, so start there rather than in the workbook. |
| **exit 2** | a gate stopped it. The question was not asked; fix what it names and re-run. |

## It cleans up after itself

Every setup gets a working tree holding a copied channel zip, a setup, a
generated linelist and a fake dataset — tens of megabytes apiece, six setups a
run. So:

- a **green** run deletes every generated tree and keeps the logs, because the
  next run rebuilds them and the logs are the record that it was green
- a **failed** run keeps everything, because what a broken run left behind is
  the only account of what went wrong
- `--keep` keeps both

## The setups

`.obt/draft/headless/setups/current/` by default — six as of 2026-11-25:
cholera-ocs, cholera, demo, ebola, gtfcc, measles. `OBT_PACKAGE_SETUPS` points
somewhere else. `--list` names them without running anything.

## Not yet exercised

**The gates are tested; the three stages are not.** As of 2026-11-25 this
machine fails gate 1 (obt 0.0.1 against 0.0.2 in source), so no setup has been
driven through `obt_designer_generate`, `obt_fake` and `obt_linelist_import` by
this runner yet. The recipe follows obt's documented API and the pattern the
removed `verify.sh` case 020 used for the generate half, but the fake-and-import
half has never run here.

Install the current obt, build a `latest` channel, and the first run is as much
a test of this suite as of the package. Expect to fix something in
`run-setup.R`.

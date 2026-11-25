# verify-package — the package suite

Drives every headless setup through the **obt package**: generates its linelist,
makes a fake dataset out of the same setup, and imports that dataset into the
linelist it just generated.

```
bash scripts/verify-package/verify-package.sh              every setup
bash scripts/verify-package/verify-package.sh --list       name them, run nothing
bash scripts/verify-package/verify-package.sh --setup ebola
bash scripts/verify-package/verify-package.sh --records 50 rows per worksheet
bash scripts/verify-package/verify-package.sh --keep       keep what it generated
bash scripts/verify-package/verify-package.sh --help       every option
```

**Close Excel first.** Each setup launches its own through obt, and the watchdog
will not start while another Excel is open. Then read the watchdog lines the run
prints at the end: `samples: 1` means the guard tripped on its first look and
exited, and everything after it ran unwatched.

## What each thing answers

| | |
|---|---|
| `scripts/tests/run-tests.R` | the **unit** suites, inside Excel, over `src/` |
| `scripts/release/preflight.sh` | the **static** gate |
| `scripts/verify/verify.sh` | the product's **behaviour**, over `src/` |
| `scripts/verify-package/` | **this**: the shipped **package** over the shipped **binaries** |

`verify.sh` builds with the mock designer carrying `src/`, so it judges the
working tree. This judges what a user installs.

## Two gates, and both must pass

Either one failing exits **2** and runs nothing, because the question was not
asked. Each prints the reason and the command that fixes it.

1. **The installed obt must be the latest.** `gates/package-version.R` compares
   the installed package against its source checkout (`OBT_SOURCE`, else the
   sibling `../obt`) by `DESCRIPTION` Version and by the installed `RemoteSha`
   against the source `HEAD`. A dirty source tree also fails, since an install
   comes from a commit. No checkout to compare against answers `unknown`, which
   also stops the run.

2. **The channel must be one obt can drive.** `gates/channel-age.R` reads the
   version off the designer's filename and applies obt's own constants through
   `obt:::`. A date-stamped zip from the old scheme is refused by obt itself.
   Build a current one with `scripts/release/build-release-zip.sh dev latest`.

The channel's **age** is printed on every result line and never gates: a dev
channel is stamped `latest` however old it is, so the version carries no age.

## Layout

| File | What it is |
|---|---|
| `verify-package.sh` | the runner: gates, the loop over setups, the table, the cleanup |
| `run-setup.R` | one setup all the way through obt; prints `PKG_*` markers |
| `gates/package-version.R` | is the installed obt the latest? |
| `gates/channel-age.R` | can obt drive the cached channel, and how old is it? |

Assertions, `inspect_linelist` and the Excel helpers come from
`scripts/verify/lib/verify-lib.sh`, so "ok" and "BAD" mean the same in both
suites.

## It cleans up after itself

A per-setup working tree holds a copied channel zip, a setup, a linelist and a
fake dataset. A **green** run deletes every tree and keeps the logs; a **failed**
run keeps everything; `--keep` keeps both.

## Not yet exercised

The gates are tested. The three stages are not: as of 2026-11-25 this machine
fails gate 1, so nothing has yet been driven through `obt_designer_generate`,
`obt_fake` and `obt_linelist_import` by this runner. Expect the first real run
to need a fix in `run-setup.R`.

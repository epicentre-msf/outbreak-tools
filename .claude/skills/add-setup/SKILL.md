---
name: add-setup
description: Take a setup workbook out of the Downloads folder and add it to the headless build — classify it as old or current with no Excel, name it, record it in the manifest, convert it if it is old, generate once to prove it, and push the folder. Use when the owner says add this setup, I downloaded a setup, put this setup in the headless folder, or test the build against a new setup.
---

# add-setup — one setup from Downloads into the headless build

The headless folder is where every setup the project can build from lives. This
is how one more gets in, and the order matters: a setup that is classified wrong
converts into nothing, and a setup named carelessly is a stem somebody has to
live with in every later batch.

Six steps, and the last two are what make it worth doing at all.

## 1. Find it, and be sure which one

```bash
ls -lt "$HOME/Downloads"/*.xls[bmx] 2>/dev/null | head -20
```

`$HOME/Downloads` is the same path on Windows Git Bash and on macOS, so nothing
here needs the owner's user name written down.

- The owner named a file → take that one.
- Exactly one workbook is newer than everything already in the folder → take it
  and say which.
- Otherwise **ask**. Adding the wrong setup costs a conversion, a build and a
  stem in the manifest that then has to be taken back out.

**A setup from Downloads is a macro workbook from outside this project.** Never
run a script from inside that folder and never point an interpreter at it as a
working directory; pass the path as an argument, which is what every command
below does. `convert-setup.ps1` opens the source with `EnableEvents = False`
precisely so its own `Workbook_Open` does not run.

**Is it a filled setup at all?** A name like `setup_dev-latest.xlsb` or
`setup_main-2.0.0.xlsb` is a *release* artifact — the empty setup. That belongs
in `template/setup-empty-current.<ext>` as the conversion target, not in
`setups/`, and no linelist is generated from it because its dictionary is empty.
Ask when the name looks like a release and the size is near the empty one
(about 2.1 MB against 2.3 MB filled).

## 2. Is it already here?

A duplicate under a second name doubles every later batch and reports two
verdicts for one setup.

```bash
# macOS has shasum, Git Bash has sha256sum; neither has both.
h() { if command -v sha256sum >/dev/null 2>&1; then sha256sum "$@"; else shasum -a 256 "$@"; fi; }
cand=$(h "$HOME/Downloads/<file>" | cut -d' ' -f1)
h .obt/draft/headless/setups/*/*.xls[bm] | grep "^$cand" || echo "not in the folder yet"
```

A line printed names the setup that already holds this content, under whatever
stem it was given. Stop there and say which one; `manifest.csv` records where
each file came from, so the original name is recoverable.

## 3. Old or current, with no Excel

The two structures are told apart by the setup's own named ranges, read straight
out of the zip in milliseconds:

| Structure | Carries |
|---|---|
| **current** | `RNG_tab_<table>_<field>_updated`, `__OBT__SILENT_OPERATIONS__` |
| **old** | `RNG_<field>_<table>` (e.g. `RNG_main_label_dict`), `RNG_DevPasswd` |

The workbook part is `xl/workbook.bin` in an `.xlsb` and `xl/workbook.xml` in an
`.xlsm`, so try both:

```bash
f="$HOME/Downloads/<file>"
for part in xl/workbook.bin xl/workbook.xml; do
  n=$(unzip -p "$f" "$part" 2>/dev/null | tr -d '\000' \
      | grep -acE 'RNG_tab_[A-Za-z0-9_]+_updated' || true)
  [ "${n:-0}" -gt 0 ] && echo "$part: CURRENT structure" && break
  unzip -p "$f" "$part" 2>/dev/null | tr -d '\000' \
    | grep -aqE 'RNG_DevPasswd|RNG_main_label_dict' && echo "$part: OLD structure" && break
done
```

Neither answer means it is not a setup. Say so rather than guessing — an export,
a linelist or a master setup all land in Downloads looking similar.

## 4. Name it, copy it, record it

The stem is what every batch, log, output folder and verdict line will carry
forever, so it is chosen deliberately: **lowercase, hyphens, no spaces, no
dates unless the date is the point, no `- Copie`, no `(1)`**. Disease, then
place or year when one is needed to tell two apart: `cholera-ocs`,
`ebola-2026-11-19`, `measles-training-2024`, `malaria-aweil`.

```bash
H=.obt/draft/headless
cp "$HOME/Downloads/<file>" "$H/setups/<current|old>/<stem>.<ext>"
```

The original stays in Downloads, untouched. Then one row in the manifest, which
is the only place the original name survives:

```bash
printf '%s,%s,%s,%s\n' "<current|old>" "<stem>.<ext>" \
  "$(wc -c < "$H/setups/<kind>/<stem>.<ext>" | tr -d ' ')" \
  "$HOME/Downloads/<file>" >> "$H/manifest.csv"
```

**A geobase, if the setup wants one.** `setups/geo-map.csv` pairs a stem with a
file under `geo/`, and a stem that is not listed builds without one, which the
build takes and reports. Add the row when the setup's geography matters, and
when the geobase itself is new, drop it in `geo/` and record it in the manifest
as `geo`.

Nothing is added to git. The whole folder is untracked by design, and the
setups are field files: `.obt/` is gitignored and the sync store is private.

## 5. Convert it, if it is old

**Windows only** — `convert-setup.ps1` drives Excel through COM. On a Mac, add
the file, push, and convert from the Windows machine.

```powershell
.obt\draft\headless\guard-start.ps1 -Tag add-setup -DeadlineSeconds 1800
.obt\draft\headless\convert-setup.ps1 `
  -Source setups\old\<stem>.xlsb -Target setups\converted\<stem>.xlsm
```

It fills a fresh copy of the current empty setup and answers `OK` or
`ERROR <n>: <text>` on its last line, with the dictionary row count above it.
**Green needs both**: `OK` and rows above zero, because the import answers `OK`
over an empty dictionary just as readily.

`ERROR 91` on its own names no sheet — the production module imports all five at
once. `-Sheets YYYYN` runs this folder's per-sheet module instead and converts
what the setup has, which is how `malaria-aweil` lands its other four sheets.
`/headless-failure` carries that measurement.

## 6. Prove it generates, then push

One build, that setup alone. Two to four minutes, Excel visible by design:

```bash
bash .obt/draft/headless/run-batch.sh <stem>               # a current setup
bash .obt/draft/headless/run-batch.sh --converted <stem>   # one just converted
```

Judge it the way `/headless-generation` does: **green is `outcome=OK` plus the
file on disk plus sane counts.** Reported `[Error]` lines belong to the setup's
own content and are reported, not held against the build — this setup's author
is the person who needs to hear them, so name them and say which variable each
belongs to.

Then make the addition real for the other machine:

1. Add a row to `headless-progress.html`, in **The setups** section, outside the
   `RESULTS` markers — stem, bytes, where it came from. The next full pass picks
   the setup up on its own, because `run-batch.sh --all-current` and
   `--all-converted` enumerate the folder.
2. `bash scripts/headless/push-headless.sh`. Without this the setup exists on one
   laptop only, and the next pass on the other machine silently builds a shorter
   list.
3. Report: the stem, the structure, the conversion row count if it was old, the
   generation verdict with its counts, the reported errors and whose they are,
   and the geobase it was paired with.

## What this does not do

It does not run the full pass (`/headless-generation`), commit anything
(`/commit-work`), or touch `src/`. A setup that exposes a product fault is a
finding to report, not a reason to start fixing the product inside this task.

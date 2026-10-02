# Releasing OutbreakTools

How binaries, releases, and the releases page work after the move off git-tracked
binaries. **TL;DR:** binaries live in a GitHub Release asset store (not git); stable
versions are cut from `CHANGELOG.md` by CI **on `main` only**, and the development
build lives in one mutable pre-release that never gets a version number.

---

## 1. The big picture

```
  Excel edit ──▶ .mock/ ──(update_files tasks)──▶ src/bin/
                    │                                  │
                    └──────────── push-assets ─────────┘
                                       │
                          ┌────────────▼─────────────────────┐
                          │ GitHub Release: working-binaries  │  off-git asset store
                          │   working-binaries.tar.gz         │  (mutable, pre-release)
                          └────────────┬─────────────────────┘
                                       │ pull-assets (any machine / CI)
                    ┌──────────────────┴───────────────────┐
                    │                                      │
  CHANGELOG.md ─push to main─▶ release.yml     push-assets ─▶ dev-latest.yml
        │                          │                              │
        │           ┌──────────────▼────────────────┐  ┌──────────▼───────────────────┐
        │           │ v<MAJOR.MINOR.PATCH> (Latest)  │  │ dev-latest (the ONE dev       │
        │           │   OBT-main-<version>.zip       │  │   pre-release, mutable)       │
        │           │   OBT-main-latest.zip (alias)  │  │   OBT-dev-latest.zip          │
        │           │ immutable, one per version     │  │ no version, rebuilt each time │
        │           └──────────────┬────────────────┘  └──────────┬───────────────────┘
        │                          │      legacy-archive (old zips, kept forever)
        └────────────── publish.yml (build-releases.sh, Releases API) ◀──────┘
                                   ▼
                      site/releases.qmd ──▶ gh-pages (releases page)
```

Binaries are **not** committed to git: `.gitignore` covers `src/bin/`, `.mock/`, the
ribbon templates (`ribbons/_ribbontemplate_*.xlsb`), and there is no `releases/`
folder. `git history` was rewritten to purge them (`scripts/history-rewrite/`); the
ribbon-template binaries are untracked now and folded into a later rewrite pass (they
are listed in `02-rewrite.sh`'s `LIVE_PATHS`). The ribbon **sources** stay in git:
one folder per distinct ribbon under `ribbons/<name>/`, each holding `ribbon.xml`
and the `images/` it references. `scripts/devtools/ribbon-extract.sh` writes those
folders and `ribbon-pack.sh` puts them back into a workbook — see
`.obt/conventions.md` §11.1.

---

## 2. The binary asset store

A single pinned **pre-release** named `working-binaries` holds one asset,
`working-binaries.tar.gz`, containing (at repo-relative paths) `src/bin/`, `.mock/`,
and the two `ribbons/_ribbontemplate_*.xlsb`.

| Task | Command | What it does |
|------|---------|--------------|
| Publish your binaries off-git | `bash scripts/release/push-assets.sh` | runs the **preflight** below, then bundles the paths above → uploads (`--clobber`). Creates the release on first run, then triggers `dev-latest.yml` so the dev pre-release picks the new binaries up. |
| Restore binaries (new machine / CI) | `bash scripts/release/pull-assets.sh` | downloads the bundle → extracts into `src/bin/`, `.mock/`, and the ribbon templates. |

Windows: use the `.ps1` twins. Both need the `gh` CLI authenticated with write
access to the repo. Set `OBT_SKIP_DEV_REFRESH=1` to publish binaries without
republishing the dev build.

### The preflight

`push-assets.sh` (and its `.ps1` twin) runs `scripts/release/preflight.sh` first and
**uploads nothing if it fails**. It is the one place that catches a binary and its
sources disagreeing, because once the `.xlsb` is in the store nothing in git can tell
you which side was right. In order:

| # | Check | Fails when |
|---|-------|-----------|
| 1 | envelope | a bundled workbook carries a `customUI14` part (conventions.md §11) |
| 2 | plumbing | `ribbon-doctor.sh` prints a `FAIL` for a workbook that has a ribbon |
| 3 | ribbons | a workbook's `customUI` or an icon differs, **byte for byte**, from the folder under `ribbons/` that holds it |
| 4 | imports | `mock-import-drift.sh` finds a hand-imported class or module that no longer matches its `.cls`/`.bas` |
| 5 | rebuild | — writes, does not check: the `ribbons/` sources are re-read out of the workbooks |

Steps 1–4 only read. Step 5 runs only once all four are green, so an aborted push
leaves the tree exactly as it found it. Run the checks on their own with
`bash scripts/release/preflight.sh --check`, which stops after step 4.

Step 3 is where drift actually shows up, and it is byte-exact on purpose —
`.gitattributes` marks `ribbons/**/ribbon.xml` as `-text` so the file in git is the
customUI part as it sits inside the workbook. (The *report* is normalised with
`--strip-trailing-cr`, because a customUI part can mix CRLF and LF and GNU diff then
calls every line changed; when that is the only difference the check says so.) The
pairing is a table at the top of `sync-ribbons.sh` rather than a name rule, because two
workbooks share one folder: `.mock/setup_mock.xlsb` and `src/bin/setup/setup_dev.xlsb`
are two copies of the one ribbon in `ribbons/setup_mock`, and `.source` names only one
of them. There is one folder per ribbon, named after the mock that owns it, and both
workbooks carrying that ribbon are checked against it, so a dev build trailing its mock
cannot reach the store.

The promoted main builds (`designer.xlsb`, `setup.xlsb`, `msetup.xlsb`) and
`_ribbontemplate_main.xlsb` have no folder and are not compared. They are promoted from
the dev line, so they trail it by design and would fail the check every time — their
ribbon lives only in the binary. `.mock/cleanDesignerMockFile.xlsb` is left out too: it
is a stripped designer kept for rebuilds, not a shipped ribbon source.

When step 3 fails, decide which side is right and make them agree:

```sh
# the folder is ahead (a ribbon edited as text)
scripts/devtools/ribbon-pack.sh <folder> --out <workbook> --overwrite

# the workbook is ahead (a ribbon edited in Excel / OfficeRibbonX)
scripts/devtools/ribbon-extract.sh <workbook> --out-dir ribbons --overwrite
```

A dev workbook trailing its mock is not a ribbon edit at all — run the `update … dev`
tasks to recopy it.

`src/tests/.input` is bundled but exempt from every check: those are frozen fixtures,
and an old setup is *supposed* to still carry an old ribbon envelope.

**Daily flow:** edit in Excel → save to `.mock/` → run the `update … dev/main` VS Code
tasks (`update_files.R`) to populate `src/bin/` → `push-assets` to publish. On another
machine, `pull-assets` first.

---

## 3. Versions

Versions are **semantic** — `vMAJOR.MINOR.PATCH` — and they exist on `main` only.

- **MAJOR** — a linelist or setup built by the previous version stops working with the new one.
- **MINOR** — new capability that stays compatible.
- **PATCH** — fixes alone.

`v1.2.0` is the baseline: the last 2024 stable build (2024-10-19), republished under
this scheme. Everything before it lives in `legacy-archive`, which is **kept forever**.

There is no such thing as a dev version. The development build is published
continuously into one mutable pre-release, `dev-latest`, and a dev build is never cut
as a numbered pre-release.

---

## 4. Cutting a stable release

Stable releases are **driven by `CHANGELOG.md`**. To release:

1. Rename the top `## [Unreleased]` heading to the version you are shipping:
   ```markdown
   ## [2.0.0]
   ### Fixed
   - ...
   ### Added
   - ...
   ```
   A trailing date is fine (`## [2.0.0] - 2026-08-14`).
2. Make sure the asset store has the binaries you want shipped (`push-assets`).
3. Merge `dev → main` and push `main`.

`release.yml` then: parses the top semantic version + its notes → `pull-assets` →
`build-release-zip main <version>` → creates the GitHub Release (notes = the changelog
section, asset = `OBT-main-<version>.zip`, plus the `OBT-main-latest.zip` alias) →
refreshes the releases page.

It is **idempotent**: re-pushing the same version does nothing (the release already
exists); a failed run self-heals the download alias on the next run. A push to `main`
that still says `## [Unreleased]` at the top resolves to the last released version and
so cuts nothing.

> The shipped zip contains a **designer**, a **setup**, a **master setup** and a
> **ribbon template**. `main` ships `designer.xlsb`/`setup.xlsb`/`msetup.xlsb`; `dev`
> (and `hot-fixes`) ship the `_dev` builds.

### Refreshing the dev build

`dev-latest.yml` rebuilds `OBT-dev-latest.zip` from the asset store and overwrites the
asset and the notes on the existing `dev-latest` pre-release. It fires when
`push-assets.sh` publishes binaries, when `CHANGELOG.md` lands on `dev`, and from the
Actions tab. It never creates a tag or a release of its own — there is only ever the
one `dev-latest`, and its notes name the commit the build came from.

### Which file goes where (two separate uploads — don't confuse them)

```
push-assets.sh ─▶ working-binaries.tar.gz   (asset store = src/bin + .mock + ribbon templates)
  you, when binaries change                       │
                                                  │  the workflows pull + build from it
                                                  ▼
release.yml     ─▶ OBT-main-<version>.zip   (the release = designer + setup + msetup + ribbon)
  CI, on a CHANGELOG version landing on main   + OBT-main-latest.zip (stable alias)
dev-latest.yml  ─▶ OBT-dev-latest.zip       (the one dev build, overwritten in place)
```

| Upload | By | To which release | Asset(s) |
|--------|----|------------------|----------|
| working binaries | **you** — `push-assets.sh` | `working-binaries` (mutable store) | `working-binaries.tar.gz` |
| a stable version | **CI** — `release.yml` | `v<version>` (Latest, immutable) | `OBT-main-<version>.zip` **+** `OBT-main-latest.zip` |
| the dev build | **CI** — `dev-latest.yml` | `dev-latest` (mutable, unversioned) | `OBT-dev-latest.zip` |

A release **reads** the asset store; it never overwrites `working-binaries.tar.gz`.

---

## 5. Branch & tag model

| Branch | Stream | Tag | GitHub "Latest"? |
|--------|--------|-----|------------------|
| `main` | stable | `vMAJOR.MINOR.PATCH` | yes (`--latest`) |
| `dev` (default) | bleeding edge | moving `dev-latest`, no version | no (pre-release) |
| `hot-fixes` | released through the dev stream | — | — |

Every tag on the remote is one of: a `vMAJOR.MINOR.PATCH` release, `dev-latest`,
`legacy-archive`, or the `working-binaries` store. Anything else is stray and should
be deleted.

### Why three tags are parked on 2021 commits

GitHub orders the releases page by the **target commit's date** — not by when the
release was published. Left alone, `legacy-archive` and `working-binaries` (tagged on
recent commits) floated above the newest version, and a `dev-latest` that chased dev
HEAD would always outrank the newest stable release.

So the three unversioned releases are parked on early commits, which puts every
version above them and keeps the newest version on top forever:

| Release | Parked on | Date |
|---------|-----------|------|
| `dev-latest` | `1fe00304` | 2021-02-01 |
| `legacy-archive` | `7313ce70` | 2021-02-01 |
| `working-binaries` | `e2bb1f46` | 2021-01-28 (initial commit) |

None of the three is a source tag — they are download pointers, and `dev-latest`
names its real build commit in its release notes. **`dev-latest.yml` must never move
the `dev-latest` tag**; it overwrites the asset and the notes in place.

The same three also carry an **empty release name**, so GitHub falls back to showing
the bare tag instead of a bold title, and each body opens with a one-sentence summary.
Only a version release gets a title, and its title is its tag (`v2.0.0`).

Version tags are the opposite: `v<version>` is always tagged on the real `main` commit
that produced it, which is what puts a new release at the top of the list.

Stable download URLs used by the README links:
- Stable: `…/releases/latest/download/OBT-main-latest.zip`
- Dev:  `…/releases/download/dev-latest/OBT-dev-latest.zip`

**Where workflows must live:** a `push`-triggered workflow runs from the **pushed
branch's** copy, so `release.yml` must be on **`main`** to cut main releases (merge
`dev → main`). `workflow_dispatch` / `release` events run from the **default branch
(`dev`)**. A workflow does not trigger until it is committed to the relevant branch.

---

## 6. The releases page

`scripts/ci/build-releases.sh` generates `site/releases.qmd` from the **GitHub
Releases API** (`gh api …/releases`): a Download table, a Versions table, per-release
changelog notes, and a Legacy-archive table. Only tags matching
`v<MAJOR>.<MINOR>.<PATCH>` reach the Versions table, so the `working-binaries` store
and the `dev-latest` pointer never appear as versions.

`publish.yml` runs it (with `GH_TOKEN`) and publishes the Quarto site to `gh-pages`.
It triggers on push to `dev`, on release events, and via `workflow_dispatch` (which is
how `release.yml` and `dev-latest.yml` refresh the page). Test it offline with
`OUT=/tmp/r.qmd RELEASES_JSON_FILE=fixture.json bash scripts/ci/build-releases.sh`.

---

## 7. File reference

| Path | Role |
|------|------|
| `CHANGELOG.md` | release log; its top semantic version drives releases |
| `scripts/release/push-assets.{sh,ps1}` | publish working binaries to the asset store, then refresh `dev-latest` |
| `scripts/release/preflight.sh` | the gate push-assets runs first; uploads nothing if it fails |
| `scripts/release/sync-ribbons.sh` | preflight steps 3 and 5: ribbons byte-compared, then rebuilt from the workbooks |
| `scripts/release/pull-assets.{sh,ps1}` | restore working binaries from the asset store |
| `scripts/release/build-release-zip.sh` | assemble `OBT-{branch}-{version}.zip` |
| `scripts/release/backfill-legacy.sh` | one-time: upload old `releases/old/*.zip` to `legacy-archive` |
| `scripts/ci/build-releases.sh` | generate the releases page from the Releases API |
| `.github/workflows/release.yml` | changelog-driven stable release, `main` only (the engine) |
| `.github/workflows/dev-latest.yml` | rebuild the one dev pre-release in place |
| `.github/workflows/publish.yml` | build docs + releases page → gh-pages |
| `scripts/history-rewrite/` | one-time git history purge tooling (README inside) |

---

## 8. Maintenance / one-time

- **History rewrite (done):** binaries were purged from all git history via
  `scripts/history-rewrite/` (single `filter-repo` pass over all branches, then
  force-push). See its `README.md` if it ever needs repeating.
- **Reclaim GitHub-side space:** force-pushing shrinks local clones, but GitHub keeps
  the old objects until it runs `gc`. Open a **GitHub Support request** to repack the
  repo and report the new size.
- **After a history rewrite, everyone re-syncs:**
  ```sh
  git fetch origin --prune --tags --force
  git reset --hard origin/dev
  for b in main hot-fixes documentation; do git update-ref "refs/heads/$b" "refs/remotes/origin/$b"; done
  ```
  (Or a fresh `git clone`, which is cleanest.)

---

## 9. Quick reference

```sh
# get the binaries (new machine)
bash scripts/release/pull-assets.sh

# after editing binaries in Excel + update_files tasks
# (also refreshes the dev-latest pre-release)
bash scripts/release/push-assets.sh

# cut a stable release: rename "## [Unreleased]" to "## [2.0.0]" in CHANGELOG.md, then
git add CHANGELOG.md && git commit -m "Release 2.0.0" && git push origin dev
git switch main && git merge dev && git push origin main    # this is what cuts v2.0.0

# preview the releases page locally
OUT=/tmp/r.qmd bash scripts/ci/build-releases.sh && less /tmp/r.qmd
```

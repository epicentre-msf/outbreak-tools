# Releasing OutbreakTools

How binaries, releases, and the releases page work after the move off git-tracked
binaries. **TL;DR:** binaries live in a GitHub Release asset store (not git);
releases are cut automatically from `CHANGELOG.md` by CI.

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
                                       ▼
  CHANGELOG.md ──push to dev/main──▶ release.yml ──▶ pull-assets → build-release-zip
        │                                              │
        │                              ┌───────────────▼──────────────────────────┐
        │                              │ GitHub Releases (per version, immutable)  │
        │                              │  v2026.06.14      (Latest)  OBT-main-….zip │
        │                              │  v2026.06.14-dev  (pre-rel) OBT-dev-….zip  │
        │                              │  dev-latest       (moving pointer)         │
        │                              │  legacy-archive   (old zips)               │
        │                              └───────────────┬──────────────────────────┘
        └───────────────────────────── publish.yml (build-releases.sh, Releases API)
                                                       ▼
                                          site/releases.qmd ──▶ gh-pages (releases page)
```

Binaries are **never** committed to git (`.gitignore` covers `src/bin/`, `.mock/`,
and there is no `releases/` folder). `git history` was rewritten to purge all
historical binaries (`automate/history-rewrite/`).

---

## 2. The binary asset store

A single pinned **pre-release** named `working-binaries` holds one asset,
`working-binaries.tar.gz`, containing (at repo-relative paths) `src/bin/`, `.mock/`,
and the two `ribbons/_ribbontemplate_*.xlsb`.

| Task | Command | What it does |
|------|---------|--------------|
| Publish your binaries off-git | `bash automate/release/push-assets.sh` | bundles the paths above → uploads (`--clobber`). Creates the release on first run. |
| Restore binaries (new machine / CI) | `bash automate/release/pull-assets.sh` | downloads the bundle → extracts into `src/bin/`, `.mock/`, ribbons. |

Windows: use the `.ps1` twins. Both need the `gh` CLI authenticated with write
access to the repo.

**Daily flow:** edit in Excel → save to `.mock/` → run the `update … dev/main` VS Code
tasks (`update_files.R`) to populate `src/bin/` → `push-assets` to publish. On another
machine, `pull-assets` first.

---

## 3. Cutting a release

Releases are **driven by `CHANGELOG.md`** (date-based versions). To release:

1. Add a new dated heading at the top of `CHANGELOG.md`:
   ```markdown
   ## [2026.06.14]
   ### Fixed
   - ...
   ### Added
   - ...
   ```
   (Same day twice? suffix `## [2026.06.14.1]`.)
2. Make sure the asset store has the binaries you want shipped (`push-assets`).
3. Push `CHANGELOG.md`:
   - **to `dev`** → cuts `v2026.06.14-dev` (**pre-release**, tagged on the dev commit) and
     re-points the moving `dev-latest`.
   - **to `main`** → cuts `v2026.06.14` (**Latest**) with the stable `OBT-main-latest.zip`.

`release.yml` then: parses the top changelog version + its notes → `pull-assets` →
`build-release-zip` → creates the GitHub Release (notes = the changelog section,
asset = `OBT-{branch}-{version}.zip`) → refreshes the releases page.

It is **idempotent**: re-pushing the same version does nothing (the release already
exists); a failed run self-heals the download aliases on the next run.

> The shipped zip contains a **designer**, a **setup**, and a **ribbon template**.
> `main` ships `designer.xlsb`/`setup.xlsb`; `dev` (and `hot-fixes`) ship the `_dev`
> builds.

### Which file goes where (two separate uploads — don't confuse them)

```
push-assets.sh ─▶ working-binaries.tar.gz   (asset store = src/bin + .mock + ribbons)
  you, when binaries change                       │
                                                  │  release.yml pulls + builds from it
                                                  ▼
release.yml    ─▶ OBT-<branch>-<version>.zip (the release = designer + setup + ribbon)
  CI, on a CHANGELOG heading              + OBT-main-latest.zip / OBT-dev-latest.zip (stable aliases)
```

| Upload | By | To which release | Asset(s) |
|--------|----|------------------|----------|
| working binaries | **you** — `push-assets.sh` | `working-binaries` (mutable store) | `working-binaries.tar.gz` |
| a `main` version | **CI** — `release.yml` | `v<version>` (Latest) | `OBT-main-<version>.zip` **+** `OBT-main-latest.zip` |
| a `dev` version | **CI** — `release.yml` | `v<version>-dev` + `dev-latest` | `OBT-dev-<version>.zip` ; `OBT-dev-latest.zip` |

A release **reads** the asset store; it never overwrites `working-binaries.tar.gz`.

---

## 4. Branch & tag model

| Branch | Stream | Tag | GitHub "Latest"? |
|--------|--------|-----|------------------|
| `main` | stable | `vYYYY.MM.DD` | yes (`--latest`) |
| `dev` (default) | bleeding edge | `vYYYY.MM.DD-dev` + moving `dev-latest` | no (pre-release) |
| `hot-fixes` | released through the dev stream | — | — |

Stable download URLs used by the README badges:
- Main: `…/releases/latest/download/OBT-main-latest.zip`
- Dev:  `…/releases/download/dev-latest/OBT-dev-latest.zip`

**Where workflows must live:** a `push`-triggered workflow runs from the **pushed
branch's** copy, so `release.yml` must be on **`main`** to cut main releases (merge
`dev → main`). `workflow_dispatch` / `release` events run from the **default branch
(`dev`)**. A workflow does not trigger until it is committed to the relevant branch.

---

## 5. The releases page

`automate/codes/build-releases.sh` generates `site/releases.qmd` from the **GitHub
Releases API** (`gh api …/releases`): a Latest section, an "All releases" table,
per-release changelog notes, and a Legacy-archive table. The `working-binaries` infra
release and the `dev-latest` pointer are excluded from the listing.

`publish.yml` runs it (with `GH_TOKEN`) and publishes the Quarto site to `gh-pages`.
It triggers on push to `dev`, on release events, and via `workflow_dispatch` (which is
how `release.yml` refreshes the page after creating a release). Test it offline with
`OUT=/tmp/r.qmd RELEASES_JSON_FILE=fixture.json bash automate/codes/build-releases.sh`.

---

## 6. File reference

| Path | Role |
|------|------|
| `CHANGELOG.md` | date-based release log; its top heading drives releases |
| `automate/release/push-assets.{sh,ps1}` | publish working binaries to the asset store |
| `automate/release/pull-assets.{sh,ps1}` | restore working binaries from the asset store |
| `automate/release/build-release-zip.sh` | assemble `OBT-{branch}-{version}.zip` |
| `automate/release/backfill-legacy.sh` | one-time: upload old `releases/old/*.zip` to `legacy-archive` |
| `automate/codes/build-releases.sh` | generate the releases page from the Releases API |
| `.github/workflows/release.yml` | changelog-driven release (the engine) |
| `.github/workflows/publish.yml` | build docs + releases page → gh-pages |
| `automate/history-rewrite/` | one-time git history purge tooling (README inside) |

---

## 7. Maintenance / one-time

- **History rewrite (done):** binaries were purged from all git history via
  `automate/history-rewrite/` (single `filter-repo` pass over all branches, then
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

## 8. Quick reference

```sh
# get the binaries (new machine)
bash automate/release/pull-assets.sh

# after editing binaries in Excel + update_files tasks
bash automate/release/push-assets.sh

# cut a release: add a "## [YYYY.MM.DD]" heading to CHANGELOG.md, then
git add CHANGELOG.md && git commit -m "Release YYYY.MM.DD" && git push origin dev   # pre-release
#   ... merge dev -> main and push main for the stable "Latest" release

# preview the releases page locally
OUT=/tmp/r.qmd bash automate/codes/build-releases.sh && less /tmp/r.qmd
```

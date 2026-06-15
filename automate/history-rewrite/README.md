# History rewrite — purge dead binaries from git history

**Status: REVIEW ONLY. Nothing here has been run. These scripts are destructive
when run and end in a force-push. Read this file fully first.**

## Why

`.git` is ~1.1 GB. Most of that is blobs from folders/files that were **already
deleted** from the working tree long ago (`misc/`, `Rscripts/`, old root `*.xlsb`,
`input/`, `output/`, `automation/`, `fakedata/`). They still live in history, so every
clone pays for them. The only way to reclaim that space is to rewrite history and
force-push.

## The dev/main consistency guarantee (the important part)

The merge conflicts you want to avoid happen when `dev` and `main` are rewritten in
**separate** passes: commits they share get **different** new SHAs on each branch, so
the branches look unrelated and merging them explodes.

We avoid that by construction:

- We run **one** `git filter-repo` pass on a **single mirror clone that contains all
  branches**. filter-repo walks the whole commit graph once and maps each original
  commit to exactly one rewritten commit. Commits shared across branches are rewritten
  **once**, so they stay shared (with new SHAs).
- Verified today: `dev..main = 0`, `main..dev = 303`, and `main`, `hot-fixes`,
  `documentation` are **all ancestors of `dev`**. After a single-pass rewrite they stay
  ancestors, so `git merge <branch>` into `dev` is a clean fast-forward / no-op.
  `02-rewrite.sh` asserts this for all three (`git merge-base --is-ancestor`) and
  refuses to declare success otherwise.
- For the **dead-only** purge, none of the purged paths exist in HEAD, so each
  branch's HEAD **tree SHA is unchanged** — proof that no live file was touched.
  `02-rewrite.sh` asserts `new_tree == old_tree` per branch.

## Scope of the purge

`02-rewrite.sh` has two path groups:

- `DEAD_PATHS` — `misc Rscripts input output automation fakedata` + 5 dead root
  `*.xlsb`. Each was verified to (a) be absent from HEAD and (b) carry real bytes on
  the pushed branches. Safe to purge **any time**, zero working-tree impact. This is
  the bulk of the reclaimable space.
  - Dropped from an earlier draft: `reference getting_started how_to dev site_libs` —
    verified **0 objects on the pushed branches** (pure no-ops). Re-check
    `01-analyze.sh`'s report before re-adding anything.
- `LIVE_PATHS` (`src/bin`, `.mock`, `releases`) — currently tracked. Purged **only**
  when `INCLUDE_LIVE_BINARIES=yes` **and** `LIVE_PURGE_ACK=yes`, and only **after**
  Workstream A has moved those binaries into the GitHub Release asset store (otherwise
  you lose them). `02-rewrite.sh` prints the exact file list it would destroy and
  refuses without the ack.
  - Caveat: `--path src/bin` is a prefix match, so it also removes
    `src/bin/test-files/*` (the `unit_tests*.xlsb` fixtures). Confirm those are in the
    asset store too, or narrow `LIVE_PATHS` to `src/bin/designer src/bin/setup
    src/bin/master-setup` first.

Sequencing (confirmed 2026-06-14): run **dead-only first** (immediate win, lowest
risk — HEAD trees provably unchanged), then fold the live paths into a second pass
after the asset migration. Each pass = one force-push + one "everyone re-syncs".

## gh-pages

`gh-pages` is **generated** by the Quarto publish workflow. The mirror contains it, so
filter-repo *does* rewrite it inside the mirror — but `03-push.sh` pushes **only**
`main/dev/hot-fixes/documentation` (+ tags), so the rewritten gh-pages is **discarded**
and origin's gh-pages is left as-is (and regenerated on the next publish). Consequence:
any bloat that lives **only** on gh-pages is **not** reclaimed by this rewrite, and the
un-pushed gh-pages ref keeps those old objects reachable on the server until it is
itself rebuilt/cleaned. Verified: none of the heavy dead paths live on gh-pages, so
this does not reduce the reclaim from the pushed branches.

## GitHub-side reclamation caveat (read this)

Force-pushing rewritten branches shrinks **your local clone** immediately. It does
**not** immediately shrink the repo on GitHub: old objects stay alive via `refs/pull/*`
(closed/open PRs), forks, un-pushed branches (incl. gh-pages), and reflogs until GitHub
runs `git gc`. GitHub does not gc on demand — **open a GitHub Support request** ("we
rewrote history, please run gc / repack and report the new size") to actually reclaim
the space. Until then the reported repo size may not drop.

## Procedure

1. Install filter-repo: `brew install git-filter-repo` (or `pipx install git-filter-repo`).
2. Push any local work first — the rewrite operates on `origin`'s history.
   Tell collaborators a rewrite is coming.
3. `bash 01-analyze.sh` — clones a mirror and prints filter-repo's size-by-path
   report. Use it to confirm/extend `DEAD_PATHS`.
4. `bash 02-rewrite.sh` — rewrites the mirror, runs verification (tree-unchanged,
   every branch still an ancestor of dev, size before/after). **Does not push.**
   (For the live purge: `INCLUDE_LIVE_BINARIES=yes LIVE_PURGE_ACK=yes bash 02-rewrite.sh`.)
5. `CONFIRM=yes bash 03-push.sh` — after you're satisfied, force-pushes the rewritten
   branches and tags.
6. Everyone else re-syncs (below) or re-clones.
7. Open the GitHub Support request to reclaim server-side space.

All three scripts anchor the mirror to an absolute path next to themselves, so they
work regardless of the directory you launch them from (just run them in order).

## Recovery for collaborators after the force-push

A plain `git fetch` will not move existing local branches or tags. Each collaborator:

```sh
git fetch origin --prune --tags --force
for b in dev main hot-fixes documentation; do
  git rev-parse --verify "$b" >/dev/null 2>&1 \
    && git update-ref "refs/heads/$b" "refs/remotes/origin/$b"
done
# rebuild any local work branch onto the new history via cherry-pick
```

Cleanest is a fresh `git clone`.

# Changelog

All notable changes to OutbreakTools are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com). Versions are
**date-based** (`[YYYY.MM.DD]`). The **topmost dated version** is what CI releases when
this file is pushed to `dev`/`main` (see `.github/workflows/release.yml`, planned —
push to `main` → Latest, push to `dev` → pre-release). `[Unreleased]` collects work not
yet cut into a release; if two releases land the same day, suffix with `.1`
(`[2026.06.14.1]`).

## [2026.06.14]
### Added
- New release workflow, documented in `RELEASING.md`: releases are now cut automatically from this changelog (`main` → Latest, `dev` → pre-release).
### Changed
- Binaries (designer, setup, ribbon templates) are no longer kept in git — they now live on your machine and in a release asset store. Pull them before you start working, push them when you're done.
- Past releases stay available on the releases page.
### Removed
- Purged all binaries from the git history, shrinking the repository and making clones much faster.

## [2026.06.11]
### Fixed
- Fixed geo drift on the `main` release (hot fix).

## [2026.03.17]
- Tagged release `v2026.03.17`.

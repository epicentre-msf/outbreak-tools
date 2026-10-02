# Changelog

All notable changes to OutbreakTools are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com). Versions are
[semantic](https://semver.org): `[MAJOR.MINOR.PATCH]`. The **topmost semantic version**
is what CI releases when this file is pushed to `main` (see
`.github/workflows/release.yml`). Versions exist on `main` only — the development
build is published continuously as the single `dev-latest` pre-release and never
carries a version of its own. `[Unreleased]` collects work not yet cut into a release.

Bump **MAJOR** when a linelist or setup built by the previous version stops working
with the new one, **MINOR** for new capability that stays compatible, **PATCH** for
fixes alone.

## [Unreleased]

<!-- Ships as 2.0.0: rename this heading to "## [2.0.0]" to cut it on main. -->

### Added
- New release workflow, documented in `RELEASING.md`: releases are cut automatically from this changelog when it lands on `main`.

### Changed
- Versions are semantic (`vMAJOR.MINOR.PATCH`) instead of date-based, and are cut on `main` only.
- The development build lives in one mutable pre-release, `dev-latest`, refreshed whenever the binaries are published. Dev builds no longer create a pre-release each.
- Binaries (designer, setup, ribbon templates) are no longer kept in git — they now live on your machine and in a release asset store. Pull them before you start working, push them when you're done.

### Fixed
- Fixed geo drift on the `main` release (hot fix).

### Removed
- Purged all binaries from the git history, shrinking the repository and making clones much faster.

## [1.2.0]

Baseline of the semantic versioning scheme: the last 2024 stable build, made
2024-10-19. Everything released before it stays in the
[legacy archive](https://github.com/epicentre-msf/outbreak-tools/releases/tag/legacy-archive).

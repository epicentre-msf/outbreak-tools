#!/usr/bin/env Rscript
#
# narrow-registry.R -- cut the test registry down to one suite folder.
#
#   Rscript scripts/devtools/narrow-registry.R <folder>[,<folder>...] [registry.yml]
#
# Several folders can be named at once, separated by commas. That is what
# reproduces a failure the folders do not show one at a time: a suite that
# leaves state behind only reaches its victim when both are in the same run.
#
# A session is verified by a probe: `src/tests/test-registry.yml` narrowed to
# the work of that session, then the usual `run-tests.R --build`. Doing that cut
# by hand is where it goes wrong -- comment a `- module:` row and leave its
# `covers:` continuation standing and the YAML dies on a duplicate key -- so
# this script makes the cut instead.
#
# WHAT IT COMMENTS OUT
# -----------------------------------------------------------------------------
# Every `- module:` row under a `tests:` key, together with its `covers:`
# continuation line, for every folder except the target. Nothing else is
# touched: each `classes:`, `modules:` and `fixtures:` row stays exactly where
# it was, because the compile closure has to stay whole or the project does not
# build and the run answers the opaque -50.
#
# The whole `helpers` block is always kept. Every CustomTest module calls
# TestHelpersLite, and the fixture modules beside it build the sheets the
# suites read.
#
# The edit is line by line and adds a single `#` in front of a kept line, so the
# file comes back byte-identical under `git checkout --`. A NARROWED REGISTRY IS
# NEVER COMMITTED. The rule is `.obt/conventions.md` section 13.
#
# Run `probe-folders.sh` beside this file to walk every folder in turn.
# =============================================================================

args <- commandArgs(trailingOnly = TRUE)
if (length(args) < 1L) {
  stop("usage: narrow-registry.R <folder>[,<folder>...] [registry.yml]", call. = FALSE)
}
targets <- trimws(strsplit(args[1], ",", fixed = TRUE)[[1]])
targets <- targets[nzchar(targets)]
path <- if (length(args) >= 2L) args[2] else "src/tests/test-registry.yml"
if (!file.exists(path)) stop("no registry at ", path, call. = FALSE)

lines <- readLines(path, warn = FALSE)

folder <- NA_character_
key <- NA_character_
kept <- character(0)
seen <- character(0)

for (i in seq_along(lines)) {
  line <- lines[i]

  m <- regmatches(line, regexec("^  - folder: (\\S+)\\s*$", line))[[1]]
  if (length(m)) {
    folder <- m[2]
    key <- NA_character_
    seen <- c(seen, folder)
    next
  }

  m <- regmatches(line, regexec("^    (classes|modules|tests|fixtures):\\s*$", line))[[1]]
  if (length(m)) {
    key <- m[2]
    next
  }

  if (is.na(folder) || is.na(key) || key != "tests") next

  is_module <- grepl("^      - module: \\S+\\s*$", line)
  is_covers <- grepl("^        covers: ", line)
  if (!is_module && !is_covers) next

  if (folder %in% c("helpers", targets)) {
    if (is_module) kept <- c(kept, sub("^- module: ", "", trimws(line)))
  } else {
    lines[i] <- paste0("#", line)
  }
}

missing <- setdiff(targets, seen)
if (length(missing)) {
  stop("no folder `", paste(missing, collapse = "`, `"), "` in ", path,
       ". Folders: ", paste(seen, collapse = ", "), call. = FALSE)
}

writeLines(lines, path)
message(sprintf("narrow-registry.R: folder=%s, %d test module(s) kept", paste(targets, collapse = "+"), length(kept)))
for (k in kept) message("  ", k)

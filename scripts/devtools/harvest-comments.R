#!/usr/bin/env Rscript
#
# harvest-comments.R -- lift every comment out of the VBA source into one text
# file that can be edited by hand and written straight back.
#
# The comments in this repository carry most of the explanation: what a class
# holds, which dictionary column says what, why an approach was dropped. Reading
# and reworking that prose 250 files at a time is the job. This script collects
# it into a single file, each block labelled with the file and the lines it came
# from, and update-comments.R puts the edited prose back where it was found.
#
#   Rscript scripts/devtools/harvest-comments.R
#   <edit all-comments.txt by hand>
#   Rscript scripts/devtools/update-comments.R --input all-comments.txt
#
# WHAT COUNTS AS A COMMENT
# -----------------------------------------------------------------------------
# A run of consecutive lines whose first non-blank character is an apostrophe.
# One long class header is therefore one block, and a bare apostrophe inside it
# keeps it whole. A comment sitting after code on the same line stays where it
# is: it belongs to that statement, and the prose worth a hand pass is never
# there.
#
# WHAT THE EDITOR SEES
# -----------------------------------------------------------------------------
# The apostrophe and the indentation are stripped, so what is left is prose. The
# writer puts them back, using the indentation each line already had. Details of
# the round trip are written into the head of the generated file, where they are
# in front of whoever is editing it.
#
# WHICH FILES
# -----------------------------------------------------------------------------
# Every .cls and .bas under src/, minus whatever src/.harvestignore lists --
# today that is the stale/ folders, dead code whose comments describe interfaces
# that no longer exist. Nine .frm files live inside the workbooks and are not in
# git, so the code behind the forms is out of reach here.
#
# Usage:
#   Rscript scripts/devtools/harvest-comments.R
#   Rscript scripts/devtools/harvest-comments.R --out /tmp/pass2.txt
#   Rscript scripts/devtools/harvest-comments.R --path src/classes/showhide
#   Rscript scripts/devtools/harvest-comments.R --path src --path scripts/headless/vba
#   Rscript scripts/devtools/harvest-comments.R --min-lines 4   # skip one-liners
#   Rscript scripts/devtools/harvest-comments.R --no-ignore     # stale/ included
#   Rscript scripts/devtools/harvest-comments.R --quiet
#
# Exit code 0 when the file is written, 1 on a bad argument or an unreadable
# path.

args <- commandArgs(trailingOnly = TRUE)
quiet <- "--quiet" %in% args
no_ignore <- "--no-ignore" %in% args

repo_root <- normalizePath(file.path(dirname(sub("^--file=", "", grep(
  "^--file=", commandArgs(trailingOnly = FALSE), value = TRUE
)[1])), "..", ".."))
setwd(repo_root)
source(file.path(repo_root, "scripts", "devtools", "comments-common.R"))

# ---------------------------------------------------------------------------
# Arguments
# ---------------------------------------------------------------------------

# Both --flag value and --flag=value are accepted. A flag given more than once
# collects every value, which is how --path narrows to several folders at once.
arg_values <- function(name) {
  values <- character(0)
  for (hit in which(args == name)) {
    if (hit == length(args)) {
      stop(sprintf("%s needs a value", name), call. = FALSE)
    }
    values <- c(values, args[hit + 1])
  }
  joined <- grep(paste0("^", name, "="), args, value = TRUE)
  c(values, sub(paste0("^", name, "="), "", joined))
}

arg_value <- function(name, default) {
  values <- arg_values(name)
  if (!length(values)) default else values[length(values)]
}

known_flags <- c("--quiet", "--no-ignore", "--out", "--path", "--min-lines")
loose <- args[startsWith(args, "--") & !args %in% known_flags &
  !grepl(paste0("^(", paste(known_flags, collapse = "|"), ")="), args)]
if (length(loose)) {
  stop(sprintf("unknown option: %s", paste(loose, collapse = " ")), call. = FALSE)
}

out_path <- arg_value("--out", "all-comments.txt")
paths <- arg_values("--path")
if (!length(paths)) paths <- "src"

min_lines <- suppressWarnings(as.integer(arg_value("--min-lines", "1")))
if (is.na(min_lines) || min_lines < 1) {
  stop("--min-lines needs a whole number of 1 or more", call. = FALSE)
}

# ---------------------------------------------------------------------------
# Scope
# ---------------------------------------------------------------------------

sources <- collect_sources(paths)
skipped <- 0

if (!no_ignore) {
  patterns <- read_ignore_patterns(file.path("src", ".harvestignore"))
  if (length(patterns)) {
    drop <- is_ignored(sources, patterns)
    skipped <- sum(drop)
    sources <- sources[!drop]
  }
}

if (!length(sources)) {
  stop("no .cls or .bas files in scope", call. = FALSE)
}

# ---------------------------------------------------------------------------
# Harvest
# ---------------------------------------------------------------------------

out <- character(0)
total_blocks <- 0
total_lines <- 0
files_with_comments <- 0

for (path in sources) {
  source_file <- read_source(path)
  lines <- source_file$lines
  blocks <- comment_blocks(lines)

  if (min_lines > 1) {
    blocks <- Filter(function(b) b$end - b$start + 1 >= min_lines, blocks)
  }
  if (!length(blocks)) {
    next
  }

  chunk <- character(0)
  for (block in blocks) {
    body <- strip_quote(lines[block$start:block$end])

    collision <- looks_like_delimiter(body)
    if (any(collision)) {
      stop(sprintf(
        paste(
          "%s line %d reads as an all-comments.txt delimiter.",
          "Reword it, or change the delimiters in comments-common.R."
        ),
        path, block$start + which(collision)[1] - 1
      ), call. = FALSE)
    }

    chunk <- c(chunk, block_delimiter(path, block$start, block$end), body)
    total_lines <- total_lines + length(body)
  }

  out <- c(
    out,
    file_record(path, md5_of(path), length(blocks)),
    chunk
  )
  total_blocks <- total_blocks + length(blocks)
  files_with_comments <- files_with_comments + 1
}

# ---------------------------------------------------------------------------
# The head of the generated file
# ---------------------------------------------------------------------------

# Read on write-back: everything above the first FILE record is skipped. It is
# written for whoever opens the file, so the rules are stated where they are
# needed rather than in this script.
preamble <- c(
  "# all-comments.txt -- every comment in the VBA source, one block at a time.",
  "#",
  sprintf(
    "# Harvested by scripts/devtools/harvest-comments.R on %s.",
    format(Sys.time(), "%Y-%m-%d %H:%M")
  ),
  sprintf(
    "# %d blocks, %d lines, %d files. Scope: %s.",
    total_blocks, total_lines, files_with_comments, paste(paths, collapse = " ")
  ),
  "#",
  "# HOW TO EDIT THIS FILE",
  "# ---------------------------------------------------------------------------",
  "# Rewrite the prose under any \"file:\" line, then write it back with:",
  "#",
  "#   Rscript scripts/devtools/update-comments.R --input <this file>",
  "#",
  "# Add --dry-run first to see what it would change and touch nothing.",
  "#",
  "# The rules the writer follows:",
  "#",
  "#   * The leading apostrophe is not here. It comes back on write-back, with",
  "#     the indentation the line already had in the source. Do not type it.",
  "#   * A blank line inside a block becomes a bare apostrophe, so the block",
  "#     stays one block and can be harvested again.",
  "#   * Lines can be added and removed freely. An added line takes the",
  "#     indentation of the block's first line.",
  "#   * A block left with no lines at all deletes those lines from the source.",
  "#   * Delete any block you are not changing. What is absent is not touched,",
  "#     and a block that comes back unchanged is skipped anyway.",
  "#   * The line numbers are read, never edited. They point at the source as it",
  "#     stood at harvest time, and the md5 on each FILE line is checked before",
  "#     anything is written, so a source file edited since the harvest is",
  "#     refused rather than mangled. Re-harvest and edit again.",
  "#   * Blocks may be reordered. Each one carries its own file name.",
  "#   * Editing an @ModuleDescription line leaves the Attribute VB_Description",
  "#     line at the top of the file behind. The writer says so; the second",
  "#     edit is by hand.",
  "#",
  "# Nothing above the first FILE line below is read back.",
  "#"
)

writeLines(c(preamble, out), out_path, useBytes = TRUE)

# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

if (!quiet) {
  folders <- dirname(sub(BLOCK_RX, "\\1", grep(BLOCK_RX, out, value = TRUE)))
  counts <- sort(table(folders), decreasing = TRUE)
  cat("blocks by folder\n")
  for (folder in names(counts)) {
    cat(sprintf("  %-45s %5d\n", folder, counts[[folder]]))
  }
  cat("\n")
  if (skipped > 0) {
    cat(sprintf("%d file(s) skipped by src/.harvestignore\n", skipped))
  }
}

cat(sprintf(
  "harvest-comments: %d blocks, %d lines, %d of %d files -> %s\n",
  total_blocks, total_lines, files_with_comments,
  length(sources), out_path
))

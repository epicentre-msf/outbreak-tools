#!/usr/bin/env Rscript
#
# update-comments.R -- write an edited all-comments.txt back into the VBA source.
#
# The other half of harvest-comments.R. Each block in the edited file carries the
# file it came from and the lines it occupied; this script puts the prose back
# there, restoring the apostrophe and the indentation that were stripped for
# editing.
#
#   Rscript scripts/devtools/harvest-comments.R
#   <edit all-comments.txt by hand>
#   Rscript scripts/devtools/update-comments.R --input all-comments.txt --dry-run
#   Rscript scripts/devtools/update-comments.R --input all-comments.txt
#
# WHAT IT REFUSES TO DO
# -----------------------------------------------------------------------------
# Line numbers go stale the moment a source file is edited by other means, and a
# stale number writes prose into the middle of code. Three checks stand in front
# of that, and any one of them failing skips the whole file rather than half of
# it:
#
#   md5        the checksum on the FILE line has to match the file on disk. A
#              source file touched since the harvest is refused outright. The fix
#              is to harvest again and edit the new file.
#   comments   every line a block claims has to still be a comment line.
#   overlap    two blocks of one file may not claim the same line.
#
# --force drops the md5 check alone. The other two always hold.
#
# WHAT IT SKIPS QUIETLY
# -----------------------------------------------------------------------------
# A block whose prose rebuilds to exactly the lines already in the file. So a
# harvest written straight back changes nothing, and a file with one edited block
# among forty is rewritten once, for that block.
#
# WHAT IT CANNOT REACH
# -----------------------------------------------------------------------------
# The Attribute VB_Description line at the top of a .cls holds a copy of the
# text in @ModuleDescription. Editing the comment leaves the attribute behind,
# and this script says so instead of guessing which of the two is wanted.
#
# Usage:
#   Rscript scripts/devtools/update-comments.R
#   Rscript scripts/devtools/update-comments.R --input /tmp/pass2.txt
#   Rscript scripts/devtools/update-comments.R --dry-run
#   Rscript scripts/devtools/update-comments.R --force     # md5 check off
#   Rscript scripts/devtools/update-comments.R --quiet
#
# Exit code 1 when any file is refused, 0 when every block was written or
# skipped.

args <- commandArgs(trailingOnly = TRUE)
quiet <- "--quiet" %in% args
dry_run <- "--dry-run" %in% args
force <- "--force" %in% args

repo_root <- normalizePath(file.path(dirname(sub("^--file=", "", grep(
  "^--file=", commandArgs(trailingOnly = FALSE), value = TRUE
)[1])), "..", ".."))
setwd(repo_root)
source(file.path(repo_root, "scripts", "devtools", "comments-common.R"))

# ---------------------------------------------------------------------------
# Arguments
# ---------------------------------------------------------------------------

arg_value <- function(name, default) {
  hits <- which(args == name)
  if (length(hits)) {
    hit <- hits[length(hits)]
    if (hit == length(args)) {
      stop(sprintf("%s needs a value", name), call. = FALSE)
    }
    return(args[hit + 1])
  }
  joined <- grep(paste0("^", name, "="), args, value = TRUE)
  if (length(joined)) {
    return(sub(paste0("^", name, "="), "", joined[length(joined)]))
  }
  default
}

known_flags <- c("--quiet", "--dry-run", "--force", "--input")
loose <- args[startsWith(args, "--") & !args %in% known_flags &
  !startsWith(args, "--input=")]
if (length(loose)) {
  stop(sprintf("unknown option: %s", paste(loose, collapse = " ")), call. = FALSE)
}

input <- arg_value("--input", "all-comments.txt")
if (!file.exists(input)) {
  stop(sprintf("no such file: %s. Run harvest-comments.R first.", input),
    call. = FALSE
  )
}

# ---------------------------------------------------------------------------
# Read the edited file
# ---------------------------------------------------------------------------

# An editor that saves CRLF is fine; the endings of the edited file mean nothing
# here, only the endings of each source file do.
input_lines <- sub("\r$", "", readLines(input, warn = FALSE), useBytes = TRUE)

checksums <- character(0)
blocks <- list()
open_block <- 0L

for (index in seq_along(input_lines)) {
  line <- input_lines[index]

  if (grepl(FILE_RECORD_RX, line, useBytes = TRUE)) {
    path <- sub(FILE_RECORD_RX, "\\1", line, useBytes = TRUE)
    checksums[path] <- sub(FILE_RECORD_RX, "\\2", line, useBytes = TRUE)
    open_block <- 0L
    next
  }

  if (grepl(BLOCK_RX, line, useBytes = TRUE)) {
    start <- as.integer(sub(BLOCK_RX, "\\2", line, useBytes = TRUE))
    span <- sub(BLOCK_RX, "\\3", line, useBytes = TRUE)
    blocks[[length(blocks) + 1L]] <- list(
      path = sub(BLOCK_RX, "\\1", line, useBytes = TRUE),
      start = start,
      end = if (nzchar(span)) as.integer(sub("^-", "", span)) else start,
      body = character(0),
      at = index
    )
    open_block <- length(blocks)
    next
  }

  # Anything before the first block is the preamble. Anything after a FILE
  # record and before its first block is nothing at all.
  if (open_block > 0L) {
    blocks[[open_block]]$body <- c(blocks[[open_block]]$body, line)
  }
}

if (!length(checksums)) {
  stop(sprintf(
    "%s holds no FILE records. It is not a harvest-comments.R file.", input
  ), call. = FALSE)
}

if (!length(blocks)) {
  cat(sprintf("update-comments: no comment blocks in %s, nothing to do\n", input))
  quit(status = 0)
}

# ---------------------------------------------------------------------------
# Write, one source file at a time
# ---------------------------------------------------------------------------

paths <- vapply(blocks, function(b) b$path, character(1))
refusals <- character(0)
reports <- character(0)
files_changed <- 0
blocks_written <- 0
lines_before <- 0
lines_after <- 0

refuse <- function(path, reason) {
  refusals <<- c(refusals, sprintf("  %-45s %s", path, reason))
}

for (path in sort(unique(paths))) {
  mine <- blocks[paths == path]

  if (!file.exists(path)) {
    refuse(path, "no such file in the repository")
    next
  }

  expected <- checksums[path]
  if (is.na(expected)) {
    refuse(path, "no FILE record carries this path, so no checksum to verify")
    next
  }
  if (!force && md5_of(path) != expected) {
    refuse(path, "changed since the harvest (md5 mismatch) -- harvest again")
    next
  }

  source_file <- read_source(path)
  lines <- source_file$lines

  # Ascending for the overlap check, descending to apply: rewriting from the
  # bottom keeps every line number above the edit valid.
  mine <- mine[order(vapply(mine, function(b) b$start, numeric(1)))]

  bad <- FALSE
  previous_end <- 0L
  for (block in mine) {
    if (block$start > block$end || block$end > length(lines)) {
      refuse(path, sprintf(
        "lines %d-%d fall outside the file (%d lines) -- harvest again",
        block$start, block$end, length(lines)
      ))
      bad <- TRUE
      break
    }
    if (block$start <= previous_end) {
      refuse(path, sprintf(
        "two blocks claim line %d (edited file line %d)", block$start, block$at
      ))
      bad <- TRUE
      break
    }
    if (!all(is_comment_line(lines[block$start:block$end]))) {
      refuse(path, sprintf(
        "lines %d-%d are no longer all comment -- harvest again",
        block$start, block$end
      ))
      bad <- TRUE
      break
    }
    previous_end <- block$end
  }
  if (bad) {
    next
  }

  changed <- 0
  before <- 0
  after <- 0
  description_edited <- FALSE

  for (block in rev(mine)) {
    original <- lines[block$start:block$end]
    rebuilt <- rebuild_block(block$body, line_indent(original))

    if (identical(rebuilt, original)) {
      next
    }

    lines <- c(
      if (block$start > 1) lines[1:(block$start - 1)] else character(0),
      rebuilt,
      if (block$end < length(lines)) {
        lines[(block$end + 1):length(lines)]
      } else {
        character(0)
      }
    )

    changed <- changed + 1
    before <- before + length(original)
    after <- after + length(rebuilt)
    if (any(grepl("@ModuleDescription", c(original, rebuilt), fixed = TRUE))) {
      description_edited <- TRUE
    }
  }

  if (changed == 0) {
    next
  }

  if (!dry_run) {
    write_source(path, lines, source_file$eol, source_file$final_newline)
  }

  files_changed <- files_changed + 1
  blocks_written <- blocks_written + changed
  lines_before <- lines_before + before
  lines_after <- lines_after + after
  reports <- c(reports, sprintf(
    "  %-45s %2d block(s), %d -> %d lines%s",
    path, changed, before, after,
    if (description_edited) "  [check Attribute VB_Description]" else ""
  ))
}

# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

if (!quiet && length(reports)) {
  cat(if (dry_run) "would change\n" else "changed\n")
  cat(reports, sep = "\n")
  cat("\n\n")
}

if (length(refusals)) {
  cat("REFUSED\n")
  cat(refusals, sep = "\n")
  cat("\n\n")
}

cat(sprintf(
  "update-comments: %d block(s) read, %d %s in %d file(s), %+d line(s), %d refused\n",
  length(blocks), blocks_written,
  if (dry_run) "would be rewritten" else "rewritten",
  files_changed, lines_after - lines_before, length(refusals)
))

if (length(refusals)) {
  quit(status = 1)
}

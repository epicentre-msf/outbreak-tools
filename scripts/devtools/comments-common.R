#
# comments-common.R -- the grammar shared by harvest-comments.R and
# update-comments.R.
#
# The two scripts are a round trip: one writes all-comments.txt, the other reads
# it back. Everything they must agree on lives here, so the format cannot drift
# apart in one script and quietly break the other -- the delimiter lines, what
# counts as a comment, how a comment line is stripped for editing and rebuilt
# afterwards, and how a file is read and written without disturbing its bytes.
#
# Nothing here prints or exits. The callers own the reporting.
#
# ENCODING AND LINE ENDINGS
# -----------------------------------------------------------------------------
# Every .cls and .bas in this repository is CRLF, and two of them hold bytes
# that are not valid UTF-8 (an accented letter inside a test string). So no
# regular expression here is allowed to look at characters: they all pass
# useBytes = TRUE, which skips the encoding check that would otherwise raise an
# error on those two files. Files are read as text and written as raw bytes with
# the ending they arrived with, so a harvest followed by an unedited write-back
# leaves the file identical to the byte.

# ---------------------------------------------------------------------------
# The all-comments.txt grammar
# ---------------------------------------------------------------------------

# One record per source file, carrying the checksum the writer verifies before
# it touches anything.
FILE_RECORD_RX <- "^={5,} FILE: (.+) \\| md5: ([0-9a-f]{32}) \\| blocks: ([0-9]+) ={5,}$"

# One delimiter per comment block. Single-line blocks print one number, so the
# third group is optional and carries its own leading dash -- R's default regex
# engine has no non-capturing group to hide it in.
BLOCK_RX <- "^-{5,} file: (.+) - lines: ([0-9]+)(-[0-9]+)? -{5,}$"

file_record <- function(path, md5, blocks) {
  sprintf(
    "===================== FILE: %s | md5: %s | blocks: %d =====================",
    path, md5, blocks
  )
}

block_delimiter <- function(path, start, end) {
  span <- if (start == end) as.character(start) else sprintf("%d-%d", start, end)
  sprintf("--------- file: %s - lines: %s ---------", path, span)
}

# A body line that would be read back as a delimiter would silently move a block
# to another file. It has never happened -- the guard is here so that if a
# comment is ever written that way, the harvest stops and says so instead of
# producing a file that writes to the wrong place.
looks_like_delimiter <- function(lines) {
  grepl(FILE_RECORD_RX, lines, useBytes = TRUE) |
    grepl(BLOCK_RX, lines, useBytes = TRUE)
}

# ---------------------------------------------------------------------------
# Comments in VBA source
# ---------------------------------------------------------------------------

# A harvested comment is a line whose first non-blank character is an
# apostrophe. A trailing comment sitting after code on the same line is left
# alone: it belongs to the statement it annotates, and the prose worth editing
# by hand is never there.
COMMENT_RX <- "^[ \t]*'"

is_comment_line <- function(lines) grepl(COMMENT_RX, lines, useBytes = TRUE)

# Contiguous runs of comment lines, as a list of start/end line numbers. A blank
# line ends a block; a bare apostrophe does not, which is why the long headers
# in this repository come out as one block each.
comment_blocks <- function(lines) {
  flags <- is_comment_line(lines)
  if (!any(flags)) {
    return(list())
  }
  runs <- rle(flags)
  ends <- cumsum(runs$lengths)
  starts <- ends - runs$lengths + 1
  lapply(which(runs$values), function(i) {
    list(start = starts[i], end = ends[i])
  })
}

# What the editor sees: the indentation and the apostrophe removed, everything
# after the apostrophe kept exactly. A line holding nothing but an apostrophe
# becomes an empty line.
strip_quote <- function(lines) sub(COMMENT_RX, "", lines, useBytes = TRUE)

# The whitespace in front of the apostrophe.
line_indent <- function(lines) {
  sub("^([ \t]*)'.*$", "\\1", lines, useBytes = TRUE)
}

# Rebuild source lines from edited prose.
#
# `indents` holds the indentation of the lines being replaced. An edited line
# keeps the indentation of the line it stands in for, so a block whose lines are
# indented differently from each other survives an edit untouched. A line the
# editor added takes the indentation of the block's first line.
#
# An empty or blank-only body line becomes a bare apostrophe rather than a blank
# line, so the block stays a single block and can be harvested again.
rebuild_block <- function(body, indents) {
  if (!length(body)) {
    return(character(0))
  }
  base <- if (length(indents)) indents[1] else ""
  own <- if (length(body) <= length(indents)) {
    indents[seq_along(body)]
  } else {
    c(indents, rep(base, length(body) - length(indents)))
  }
  text <- ifelse(grepl("^[ \t]*$", body, useBytes = TRUE), "", body)
  paste0(own, "'", text)
}

# ---------------------------------------------------------------------------
# Reading and writing source files
# ---------------------------------------------------------------------------

# readLines drops CRLF on the way in, so the ending has to be read from the raw
# bytes and put back on the way out. Whether the last line ends with a newline
# is recorded for the same reason.
read_source <- function(path) {
  size <- file.info(path)$size
  raw <- if (is.na(size) || size == 0) raw(0) else readBin(path, "raw", size)
  lines <- readLines(path, warn = FALSE)
  list(
    lines = lines,
    eol = if (any(raw == as.raw(13L))) "\r\n" else "\n",
    final_newline = length(raw) == 0 || raw[length(raw)] == as.raw(10L)
  )
}

write_source <- function(path, lines, eol, final_newline) {
  text <- paste0(
    paste(lines, collapse = eol),
    if (final_newline && length(lines)) eol else ""
  )
  con <- file(path, open = "wb")
  on.exit(close(con), add = TRUE)
  writeBin(charToRaw(text), con)
  invisible(TRUE)
}

md5_of <- function(path) unname(tools::md5sum(path))

# ---------------------------------------------------------------------------
# Which files are in scope
# ---------------------------------------------------------------------------

# src/.harvestignore, in the three pattern forms .docignore uses: a trailing "/"
# is a folder anywhere in the path, a pattern holding * or ? is matched against
# the file name, anything else is a substring of the path.
read_ignore_patterns <- function(path) {
  if (!file.exists(path)) {
    return(character(0))
  }
  patterns <- trimws(readLines(path, warn = FALSE))
  patterns[nzchar(patterns) & !startsWith(patterns, "#")]
}

is_ignored <- function(rel_paths, patterns) {
  drop <- rep(FALSE, length(rel_paths))
  for (pattern in patterns) {
    hit <- if (endsWith(pattern, "/")) {
      grepl(pattern, rel_paths, fixed = TRUE)
    } else if (grepl("[*?]", pattern)) {
      grepl(utils::glob2rx(pattern), basename(rel_paths), useBytes = TRUE)
    } else {
      grepl(pattern, rel_paths, fixed = TRUE)
    }
    drop <- drop | hit
  }
  drop
}

# Every .cls and .bas under the requested paths, relative to the repo root,
# which is the working directory both callers set before they get here.
collect_sources <- function(paths) {
  found <- character(0)
  for (path in paths) {
    clean <- sub("/+$", "", path)
    if (dir.exists(clean)) {
      found <- c(found, list.files(
        clean,
        pattern = "\\.(cls|bas)$", recursive = TRUE, full.names = TRUE
      ))
    } else if (file.exists(clean) && grepl("\\.(cls|bas)$", clean)) {
      found <- c(found, clean)
    } else {
      stop(sprintf("no such folder or VBA file: %s", path), call. = FALSE)
    }
  }
  sort(unique(sub("^\\./", "", found)))
}

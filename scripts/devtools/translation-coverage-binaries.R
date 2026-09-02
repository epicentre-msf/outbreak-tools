#!/usr/bin/env Rscript
#
# translation-coverage-binaries.R -- no translation the product asks for may
# be missing from the designer binaries.
#
#   Rscript scripts/devtools/translation-coverage-binaries.R [workbook.xlsb ...]
#
# With no argument it checks the two designers a build can read:
#   .mock/designer_mock.xlsb
#   src/bin/designer/designer_dev.xlsb
#
# trads/designer_translations.xlsx is the master and translation-coverage.R
# already gates it -- but a generated linelist reads its tables from the
# DESIGNER BINARY, not from the master, so a row added to the xlsx and never
# pasted into the designer still ships a linelist whose screen reads the tag
# itself. That gap is exactly what this closes: it lifts the eight translation
# tables out of each binary and runs translation-coverage.R against them, so
# one oracle serves the master and the binaries alike.
#
# HOW THE TABLES LEAVE THE BINARY
# -----------------------------------------------------------------------------
# An .xlsb stores its sheets as BIFF12 record streams, which no reader in this
# toolchain opens. The tags and the translations are strings, so a minimal
# reader is enough: shared strings from xl/sharedStrings.bin, string cells
# (isst, inline and formula-result) from the two translation sheets, nothing
# else. The eight tables are found by their header rows -- the five language
# code columns are the invariant every lookup keys on -- and rebuilt, in
# sheet order, into a temporary .xlsx whose tables carry the master's table
# names. translation-coverage.R then reads that workbook the same way it reads
# the master. Plain R throughout: no Excel, no COM, the same on both OS.
#
# THE GATE IS MISSING, NOT DEAD
# -----------------------------------------------------------------------------
# The binaries carry the same dead rows the master carries, and a dead row
# costs upkeep, not a broken screen. A push is stopped by a MISSING tag only.
#
# Exit code: the number of workbooks with at least one missing tag.
# preflight.sh runs it before assets are pushed.

args <- commandArgs(trailingOnly = TRUE)

# Colour the way mock-import-drift.sh does: on for a terminal, off for a pipe,
# a file, NO_COLOR or --no-color, so the output stays greppable redirected.
use_color <- isatty(stdout()) && !nzchar(Sys.getenv("NO_COLOR")) &&
  !("--no-color" %in% args)
args <- args[args != "--no-color"]
C_RED <- if (use_color) "\033[1;31m" else ""
C_GREEN <- if (use_color) "\033[32m" else ""
C_OFF <- if (use_color) "\033[0m" else ""

script_path <- sub("^--file=", "",
                   grep("^--file=", commandArgs(trailingOnly = FALSE),
                        value = TRUE)[1])
repo_root <- normalizePath(file.path(dirname(script_path), "..", ".."))
setwd(repo_root)

coverage <- file.path("scripts", "devtools", "translation-coverage.R")
if (!file.exists(coverage)) {
  stop("scripts/devtools/translation-coverage.R is missing", call. = FALSE)
}
suppressMessages(library(openxlsx))

workbooks <- if (length(args)) args else c(
  ".mock/designer_mock.xlsb",
  "src/bin/designer/designer_dev.xlsb"
)

# The stacked tables of each sheet. AN UNORDERED SET on purpose: the binaries
# do not stack them in the master's order -- designer_dev holds shape, range,
# msg, drop where the master holds msg, range, shape, drop -- so each table is
# identified by its own name read out of xl/tables/*.bin, never by position.
TABLES_OF <- list(
  LinelistTranslation = c("t_tradllshapes", "t_tradllmsg",
                          "t_tradllforms", "t_tradllribbon"),
  DesignerTranslation = c("t_tradmsg", "t_tradrange",
                          "t_tradshape", "t_traddrop")
)
LANG_CODES <- c("ENG", "FRA", "SPA", "POR", "ARA")

# ---------------------------------------------------------------------------
# BIFF12 plumbing. A record is an id of one or two bytes (seven payload bits
# each, the high bit says a byte follows) and a length of up to four bytes in
# the same varint scheme, then the payload.
# ---------------------------------------------------------------------------
u32_at <- function(b, off) {
  b[off] + b[off + 1L] * 256 + b[off + 2L] * 65536 + b[off + 3L] * 16777216
}

utf16_at <- function(b, off, n_chars) {
  if (n_chars == 0) return("")
  i <- off + seq_len(n_chars) * 2L - 2L
  intToUtf8(b[i] + b[i + 1L] * 256)
}

# Walk every record, handing (id, start, size) to the callback. The bytes are
# one integer vector for the whole part; offsets are 1-based.
walk_records <- function(b, on_record) {
  pos <- 1L
  n <- length(b)
  while (pos <= n) {
    id <- b[pos]
    pos <- pos + 1L
    if (id >= 128L) {
      id <- bitwAnd(id, 127L) + bitwShiftL(bitwAnd(b[pos], 127L), 7L)
      pos <- pos + 1L
    }
    size <- 0L
    shift <- 0L
    repeat {
      byte <- b[pos]
      pos <- pos + 1L
      size <- size + bitwShiftL(bitwAnd(byte, 127L), shift)
      if (byte < 128L) break
      shift <- shift + 7L
    }
    if (pos + size - 1L > n) break
    on_record(id, pos, size)
    pos <- pos + size
  }
}

part_bytes <- function(dir, name) {
  path <- file.path(dir, name)
  if (!file.exists(path)) return(NULL)
  as.integer(readBin(path, "raw", n = file.info(path)$size))
}

# --- the shared strings: BrtSstItem (19) is flags byte + cch + utf16 --------
read_shared_strings <- function(dir) {
  b <- part_bytes(dir, "xl/sharedStrings.bin")
  if (is.null(b)) return(character(0))
  out <- character(0)
  walk_records(b, function(id, at, size) {
    if (id == 19L) {
      cch <- u32_at(b, at + 1L)
      out[[length(out) + 1L]] <<- utf16_at(b, at + 5L, cch)
    }
  })
  out
}

# --- which sheet part carries which sheet: BrtBundleSh (156) + the rels -----
sheet_parts <- function(dir) {
  b <- part_bytes(dir, "xl/workbook.bin")
  if (is.null(b)) stop("no xl/workbook.bin in this package", call. = FALSE)

  rel_of <- list()
  walk_records(b, function(id, at, size) {
    if (id != 156L) return(invisible(NULL))
    pos <- at + 8L                       # hsState + iTabID
    cch <- u32_at(b, pos)
    if (cch == 4294967295) {             # null relId
      rel_id <- ""
      pos <- pos + 4L
    } else {
      rel_id <- utf16_at(b, pos + 4L, cch)
      pos <- pos + 4L + cch * 2L
    }
    cch <- u32_at(b, pos)
    sheet_name <- utf16_at(b, pos + 4L, cch)
    rel_of[[sheet_name]] <<- rel_id
  })

  rels_path <- file.path(dir, "xl/_rels/workbook.bin.rels")
  rels <- paste(readLines(rels_path, warn = FALSE, encoding = "UTF-8"),
                collapse = "")
  targets <- list()
  for (m in regmatches(rels,
         gregexpr('<Relationship [^>]*/>', rels))[[1]]) {
    id <- sub('.*Id="([^"]*)".*', "\\1", m)
    target <- sub('.*Target="([^"]*)".*', "\\1", m)
    targets[[id]] <- target
  }

  lapply(rel_of, function(rid) {
    t <- targets[[rid]]
    if (is.null(t)) return(NULL)
    if (!grepl("^/", t)) t <- file.path("xl", t)
    sub("^/", "", t)
  })
}

# --- which table starts on which row of a sheet -----------------------------
# The sheet's rels name its table parts; each part's BrtBeginList (343) leads
# with the table's range, so its first u32 is the header row. The table's own
# name sits later in the same record behind fields this reader has no need to
# walk, so it is found instead: the record is scanned for any utf16 counted
# string reading t_trad*, which is what every translation table is called.
# A table named anything else (the EPIW* week lists) simply maps to nothing.
table_names_by_row <- function(dir, sheet_part) {
  rels_path <- file.path(dir, dirname(sheet_part), "_rels",
                         paste0(basename(sheet_part), ".rels"))
  if (!file.exists(rels_path)) return(list())
  rels <- paste(readLines(rels_path, warn = FALSE, encoding = "UTF-8"),
                collapse = "")
  targets <- regmatches(rels, gregexpr('Target="[^"]*tables/[^"]*"',
                                       rels))[[1]]
  out <- list()
  for (target in targets) {
    part <- sub('^Target="', "", sub('"$', "", target))
    part <- file.path("xl", sub("^(\\.\\./)+", "", part))
    b <- part_bytes(dir, part)
    if (is.null(b)) next
    walk_records(b, function(id, at, size) {
      if (id != 343L) return(invisible(NULL))
      header_row <- u32_at(b, at)
      name <- NULL
      for (p in seq_len(size - 12L)) {
        cch <- u32_at(b, at + p - 1L)
        if (cch < 6 || cch > 40 || at + p + 3L + cch * 2L > at + size) next
        # a table name is plain ASCII; anything else here is the scan reading
        # arbitrary bytes as a length, and must never reach intToUtf8
        i <- (at + p + 3L) + seq_len(cch) * 2L - 2L
        units <- b[i] + b[i + 1L] * 256
        if (any(units < 32 | units > 126)) next
        text <- intToUtf8(units)
        if (grepl("^t_trad", tolower(text))) {
          name <- tolower(text)
          break
        }
      }
      if (!is.null(name)) out[[as.character(header_row)]] <<- name
    })
  }
  out
}

# --- one sheet's string cells: list of (row, col, text) ---------------------
# BrtRowHdr (0) carries the row; a cell record leads with col and style, then
# its payload. Strings arrive as a shared-string index (BrtCellIsst, 7), an
# inline string (BrtCellSt, 6) or a formula's string result (BrtFmlaString, 8).
read_string_cells <- function(dir, part, shared) {
  b <- part_bytes(dir, part)
  if (is.null(b)) stop("missing sheet part ", part, call. = FALSE)
  cells <- new.env(parent = emptyenv())
  current_row <- -1L
  walk_records(b, function(id, at, size) {
    if (id == 0L) {
      current_row <<- u32_at(b, at)
      return(invisible(NULL))
    }
    if (!(id %in% c(6L, 7L, 8L))) return(invisible(NULL))
    col <- u32_at(b, at)
    text <- if (id == 7L) {
      isst <- u32_at(b, at + 8L)
      if (isst < length(shared)) shared[isst + 1L] else ""
    } else {
      cch <- u32_at(b, at + 8L)
      utf16_at(b, at + 12L, cch)
    }
    assign(sprintf("%d:%d", current_row, col), text, envir = cells)
  })
  cells
}

cell_of <- function(cells, row, col) {
  key <- sprintf("%d:%d", row, col)
  if (exists(key, envir = cells, inherits = FALSE)) {
    get(key, envir = cells, inherits = FALSE)
  } else {
    ""
  }
}

# --- the stacked tables of one sheet ----------------------------------------
# A header row is the row whose five language columns each carry one of the
# five codes; its table is the t_trad* name whose range starts on that row.
# Rows follow until the tag column goes empty. Column B (index 1) is the tag,
# C..G (2..6) the languages. Every wanted table must be found, by name -- a
# header row no table part explains, or a wanted name that never turns up,
# stops the run rather than mislabelling a table.
tables_of_sheet <- function(cells, name_by_row, wanted) {
  keys <- ls(envir = cells)
  rows <- sort(unique(as.integer(sub(":.*$", "", keys))))

  out <- list()
  for (r in rows) {
    langs <- toupper(vapply(2:6, function(cl) cell_of(cells, r, cl),
                            character(1)))
    if (!(all(langs %in% LANG_CODES) && !anyDuplicated(langs))) next

    table_name <- name_by_row[[as.character(r)]]
    if (is.null(table_name) || !(table_name %in% wanted)) next

    header <- vapply(1:6, function(cl) cell_of(cells, r, cl), character(1))
    if (!nzchar(header[1])) header[1] <- "msg_id"

    body <- list()
    at <- r + 1L
    while (nzchar(cell_of(cells, at, 1L))) {
      body[[length(body) + 1L]] <-
        vapply(1:6, function(cl) cell_of(cells, at, cl), character(1))
      at <- at + 1L
    }
    if (!length(body)) {
      stop(sprintf("table %s has a header and no rows", table_name),
           call. = FALSE)
    }
    df <- as.data.frame(do.call(rbind, body), stringsAsFactors = FALSE)
    names(df) <- header
    out[[table_name]] <- df
  }

  lost <- setdiff(wanted, names(out))
  if (length(lost)) {
    stop(sprintf("table(s) not found on the sheet: %s",
                 paste(lost, collapse = ", ")), call. = FALSE)
  }
  out
}

# ---------------------------------------------------------------------------
# One workbook: unpack, lift the tables, rebuild the xlsx, run the coverage.
# ---------------------------------------------------------------------------
check_workbook <- function(wb_path) {
  work <- file.path(tempdir(), paste0("tradbin-", basename(wb_path)))
  unlink(work, recursive = TRUE)
  dir.create(work, recursive = TRUE)
  on.exit(unlink(work, recursive = TRUE), add = TRUE)

  utils::unzip(wb_path, exdir = work)
  parts <- sheet_parts(work)
  shared <- read_shared_strings(work)

  built <- createWorkbook()
  for (sheet_name in names(TABLES_OF)) {
    part <- parts[[sheet_name]]
    if (is.null(part)) {
      stop(sprintf("%s carries no sheet named %s", wb_path, sheet_name),
           call. = FALSE)
    }
    cells <- read_string_cells(work, part, shared)
    tables <- tables_of_sheet(cells, table_names_by_row(work, part),
                              TABLES_OF[[sheet_name]])

    addWorksheet(built, sheet_name)
    at_row <- 2L
    for (table_name in names(tables)) {
      writeDataTable(built, sheet_name, tables[[table_name]],
                     startCol = 2L, startRow = at_row,
                     tableName = table_name)
      at_row <- at_row + nrow(tables[[table_name]]) + 3L
    }
  }

  lifted <- file.path(work, "lifted.xlsx")
  report <- file.path(work, "coverage.md")
  saveWorkbook(built, lifted, overwrite = TRUE)

  # No --quiet: the "missing N, dead M" line read below is printed through
  # say(), and --quiet silences it.
  said <- suppressWarnings(
    system2("Rscript", c(coverage, "--workbook", shQuote(lifted),
                         "--out", shQuote(report)),
            stdout = TRUE, stderr = TRUE)
  )

  counts <- grep("^missing [0-9]+, dead [0-9]+$", said, value = TRUE)
  if (length(counts) != 1L) {
    cat(sprintf("== %s : %stranslation-coverage.R answered no count%s\n",
                wb_path, C_RED, C_OFF))
    cat(paste0("   ", said, "\n"), sep = "")
    return(1L)
  }
  n_missing <- as.integer(sub("^missing ([0-9]+),.*$", "\\1", counts))

  if (n_missing == 0L) {
    cat(sprintf("== %s : %sno missing translation%s\n",
                wb_path, C_GREEN, C_OFF))
    return(0L)
  }

  cat(sprintf("== %s : %s%d missing translation(s)%s\n",
              wb_path, C_RED, n_missing, C_OFF))
  # The report's Missing sections name each tag; the tag rows are the ones a
  # reader acts on, so they carry the colour.
  lines <- readLines(report, warn = FALSE, encoding = "UTF-8")
  keep <- FALSE
  for (line in lines) {
    if (grepl("^## ", line)) keep <- grepl("^## Missing", line)
    if (!keep || !nzchar(line)) next
    if (grepl("`", line, fixed = TRUE) || grepl("^- ", line)) {
      cat("   ", C_RED, line, C_OFF, "\n", sep = "")
    } else {
      cat("   ", line, "\n", sep = "")
    }
  }
  1L
}

failed <- 0L
for (wb in workbooks) {
  if (!file.exists(wb)) {
    cat(sprintf("== %s : %sABSENT%s\n", wb, C_RED, C_OFF))
    failed <- failed + 1L
    next
  }
  failed <- failed + tryCatch(check_workbook(wb), error = function(e) {
    cat(sprintf("== %s : %sERROR %s%s\n",
                wb, C_RED, conditionMessage(e), C_OFF))
    1L
  })
}

tone <- if (failed > 0L) C_RED else C_GREEN
cat(sprintf("\n%sdone: %d workbook(s) with a gap%s\n", tone, failed, C_OFF))
quit(status = failed)

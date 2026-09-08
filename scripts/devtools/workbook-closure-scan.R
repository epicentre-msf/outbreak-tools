#!/usr/bin/env Rscript
#
# workbook-closure-scan.R -- are the dev binaries closed?
#
#   Rscript scripts/devtools/workbook-closure-scan.R [workbook.xlsb ...]
#
# With no argument it scans the three dev binaries:
#   src/bin/setup/setup_dev.xlsb
#   src/bin/msetup/msetup_dev.xlsb
#   src/bin/designer/designer_dev.xlsb
#
# "Closed" is three things, checked in this order for every workbook:
#
#   1. declared   every module, class and form the Dev sheet tables declare is
#                 in the VBA project, the interface of a class flagged "yes"
#                 included, and every "form modules" row has had its code
#                 copied into the document or form it names.
#   2. closure    every src component the pasted code names is in the project.
#   3. callbacks  every ribbon callback of the workbook's own customUI resolves
#                 to a Sub or Function in a standard module.
#
# WHY THREE
# -----------------------------------------------------------------------------
# The setup, master setup and designer workbooks compile from whatever somebody
# last pasted into them. A component left out is invisible until the first line
# that needs it dies with "Sub or Function not defined" -- a project-wide
# compile failure. Check 2 catches that, but it reads the pasted code as the
# ground truth, so a mock with nothing pasted into it passes: an empty
# ThisWorkbook names nothing. The Dev sheet is what says what the workbook is
# supposed to carry, and the ribbon is what says which procedures a click will
# ask for, so those two are read as well and read first. Nothing else checks
# any of this: the test harness imports its own closure from src on every run,
# so it never meets the pasted set.
#
# HOW THE WORKBOOK IS READ
# -----------------------------------------------------------------------------
# The code comes out of xl/vbaProject.bin through vba-inspect.R's dump mode,
# which decompresses every module stream in plain R, document modules included
# (ThisWorkbook carries pasted event logic). The Dev tables come out of the
# sheet cells through the readxlsb package: every sheet is read, and a table is
# found by the tag cell Development writes above it ("general modules",
# "general classes", "general form modules", "general forms"), with the header
# on the next row and the rows below. Tests tables are left alone: a dev binary
# does not ship them. The ribbon is the customUI part inside the workbook
# itself, not the copy under ribbons/. No Excel, no COM, so it runs the same on
# Windows and macOS.
#
# The candidate universe of check 2 is the file basenames of src/classes/** and
# src/modules/** minus stale/, so a name that lives inside another component --
# ProjectError inside Checking, an Office type like IRibbonControl -- is never
# a false alarm. Check 2 is a NAME scan, not a compiler: a local variable
# spelled exactly like a missing component can flag it, and a call built
# through Application.Run cannot be seen. Read the referrer list before acting.
#
# Exit code: the number of workbooks with at least one gap in any check.
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

inspect <- file.path("scripts", "devtools", "vba-inspect.R")
if (!file.exists(inspect)) {
  stop("scripts/devtools/vba-inspect.R is missing", call. = FALSE)
}
if (!requireNamespace("readxlsb", quietly = TRUE)) {
  stop("the readxlsb package is needed to read the Dev sheet tables: ",
       "install.packages(\"readxlsb\")", call. = FALSE)
}

workbooks <- if (length(args)) args else c(
  "src/bin/setup/setup_dev.xlsb",
  "src/bin/msetup/msetup_dev.xlsb",
  "src/bin/designer/designer_dev.xlsb"
)

# ---------------------------------------------------------------------------
# The candidate universe: one entry per src component, stale/ left out.
# ---------------------------------------------------------------------------
source_files <- list.files(c("src/classes", "src/modules"),
                           pattern = "\\.(cls|bas)$", recursive = TRUE,
                           full.names = TRUE)
source_files <- source_files[!grepl("/stale/", source_files, fixed = TRUE)]

component_of <- sub("\\.(cls|bas)$", "", basename(source_files))
first_hit <- !duplicated(tolower(component_of))
candidates <- setNames(dirname(source_files[first_hit]),
                       tolower(component_of[first_hit]))
candidate_names <- setNames(component_of[first_hit],
                            tolower(component_of[first_hit]))
candidate_files <- setNames(source_files[first_hit],
                            tolower(component_of[first_hit]))
cat(sprintf("candidate components in src: %d\n", length(candidates)))

# ---------------------------------------------------------------------------
# Code helpers.
# ---------------------------------------------------------------------------

# Every line of a module with its string literals and comments removed. A quote
# inside a VBA string is doubled, so the literal pattern eats whole strings
# before the apostrophe check -- "LLGeo" inside a TypeName compare never reads
# as a reference.
code_lines <- function(path) {
  lines <- readLines(path, warn = FALSE, encoding = "latin1")
  lines <- gsub('"([^"]|"")*"', '""', lines)
  sub("'.*$", "", lines)
}

# References inside one dumped module. VBA identifiers are case-insensitive,
# so the lookup is on the lowered token.
references_in <- function(path, self_name) {
  lines <- code_lines(path)
  tokens <- unlist(regmatches(lines,
                              gregexpr("[A-Za-z_][A-Za-z0-9_]*", lines)))
  tokens <- unique(tolower(tokens))
  tokens <- tokens[tokens %in% names(candidates)]
  setdiff(tokens, tolower(self_name))
}

# The procedures a module declares, lowered: Sub, Function and Property heads.
procedures_in <- function(path) {
  lines <- code_lines(path)
  heads <- regmatches(lines, regexpr(
    "^\\s*(Public\\s+|Private\\s+|Friend\\s+)?(Static\\s+)?(Sub|Function|Property\\s+(Get|Let|Set))\\s+[A-Za-z_][A-Za-z0-9_]*",
    lines, ignore.case = TRUE))
  unique(tolower(sub(".*\\s", "", heads)))
}

# ---------------------------------------------------------------------------
# The Dev tables.
#
# Development lays every table out the same way: the folder cell, the tag cell
# under it, the header row under that and the rows below, tables side by side.
# The tag cell is the invariant, so every sheet is read and every cell holding
# a tag starts a table. Rows are read down the first column to the end of the
# sheet, blanks skipped -- nothing but the table sits under a tag.
# ---------------------------------------------------------------------------
DEV_TAGS <- c("general modules", "general classes", "general form modules",
              "general forms")
TEST_TAGS <- c("tests modules", "tests classes")

sheet_count <- function(wb) {
  parts <- tryCatch(utils::unzip(wb, list = TRUE)$Name,
                    warning = function(w) character(0),
                    error = function(e) character(0))
  length(grep("^xl/worksheets/sheet[0-9]+\\.bin$", parts))
}

sheet_matrix <- function(wb, index) {
  df <- tryCatch(
    suppressWarnings(readxlsb::read_xlsb(wb, sheet = index, col_names = FALSE)),
    error = function(e) NULL
  )
  if (is.null(df) || !nrow(df) || !ncol(df)) return(NULL)
  m <- as.matrix(df)
  m[is.na(m)] <- ""
  m <- apply(m, 2, as.character)
  if (is.null(dim(m))) m <- matrix(m, nrow = 1)
  trimws(m)
}

# Answers a list of tables: tag, folder, first (column 1), second (column 2).
dev_tables <- function(wb) {
  tables <- list()
  tests_seen <- 0L
  for (index in seq_len(sheet_count(wb))) {
    m <- sheet_matrix(wb, index)
    if (is.null(m)) next
    low <- tolower(m)
    hits <- which(matrix(low %in% c(DEV_TAGS, TEST_TAGS), nrow = nrow(m)),
                  arr.ind = TRUE)
    if (!nrow(hits)) next
    for (k in seq_len(nrow(hits))) {
      r <- hits[k, 1]
      cc <- hits[k, 2]
      tag <- low[r, cc]
      if (tag %in% TEST_TAGS) {
        tests_seen <- tests_seen + 1L
        next
      }
      if (r + 2 > nrow(m)) next
      body <- seq.int(r + 2, nrow(m))
      first <- m[body, cc]
      second <- if (cc < ncol(m)) m[body, cc + 1] else rep("", length(body))
      keep <- nzchar(first)
      tables[[length(tables) + 1L]] <- list(
        tag = tag,
        folder = if (r > 1) m[r - 1, cc] else "",
        first = first[keep],
        second = second[keep]
      )
    }
  }
  attr(tables, "tests_seen") <- tests_seen
  tables
}

# ---------------------------------------------------------------------------
# The ribbon callbacks, read out of the workbook's own customUI part. A
# callback attribute is any on* or get* attribute, and its value may be
# qualified -- Module.Proc, 'Book.xlsb'!Proc -- so the trailing identifier
# is what the project has to declare.
# ---------------------------------------------------------------------------
ribbon_callbacks <- function(wb, work) {
  parts <- tryCatch(utils::unzip(wb, list = TRUE)$Name,
                    warning = function(w) character(0),
                    error = function(e) character(0))
  xml_parts <- parts[grepl("^customUI/[^/]*\\.xml$", parts)]
  if (!length(xml_parts)) return(NULL)
  callbacks <- character(0)
  for (part in xml_parts) {
    got <- tryCatch(utils::unzip(wb, files = part, exdir = work),
                    warning = function(w) character(0),
                    error = function(e) character(0))
    if (!length(got)) next
    xml <- paste(readLines(got[1], warn = FALSE, encoding = "UTF-8"),
                 collapse = " ")
    found <- regmatches(xml, gregexpr(
      "\\b(on|get)[A-Z][A-Za-z]*\\s*=\\s*\"[^\"]*\"", xml))[[1]]
    values <- sub("^[^\"]*\"", "", sub("\"$", "", found))
    values <- sub(".*[.!]", "", values)
    callbacks <- c(callbacks, values[nzchar(values)])
  }
  sort(unique(callbacks))
}

# ---------------------------------------------------------------------------
# The workbooks.
# ---------------------------------------------------------------------------
failed <- 0L

for (wb in workbooks) {
  cat("\n")
  if (!file.exists(wb)) {
    cat(sprintf("== %s : %sABSENT%s\n", wb, C_RED, C_OFF))
    failed <- failed + 1L
    next
  }

  work <- file.path(tempdir(), paste0("closure-", basename(wb)))
  unlink(work, recursive = TRUE)
  dir.create(work, recursive = TRUE)

  bin <- tryCatch(
    utils::unzip(wb, files = "xl/vbaProject.bin", exdir = work),
    warning = function(w) character(0), error = function(e) character(0)
  )
  if (!length(bin)) {
    cat(sprintf("== %s : %sno xl/vbaProject.bin (no VBA project)%s\n",
                wb, C_RED, C_OFF))
    failed <- failed + 1L
    unlink(work, recursive = TRUE)
    next
  }

  dump_dir <- file.path(work, "dump")
  rows <- suppressWarnings(
    system2("Rscript", c(inspect, shQuote(bin[1]), "dump", shQuote(dump_dir)),
            stdout = TRUE, stderr = TRUE)
  )
  parsed <- strsplit(rows[grepl("\t", rows, fixed = TRUE)], "\t", fixed = TRUE)
  if (!length(parsed)) {
    cat(sprintf("== %s : %svba-inspect.R answered nothing%s\n",
                wb, C_RED, C_OFF))
    cat(paste0("   ", rows, "\n"), sep = "")
    failed <- failed + 1L
    unlink(work, recursive = TRUE)
    next
  }

  comp_kind <- vapply(parsed, `[`, character(1), 1L)
  comp_name <- vapply(parsed, `[`, character(1), 2L)
  comp_file <- vapply(parsed, `[`, character(1), 3L)
  comp_status <- vapply(parsed, `[`, character(1), 4L)
  present <- tolower(comp_name)
  kind_of <- setNames(comp_kind, present)
  file_of <- setNames(comp_file, present)
  status_of <- setNames(comp_status, present)

  cat(sprintf("== %s : %d components\n", wb, length(comp_name)))
  undecoded <- comp_name[comp_status == "noextract"]
  if (length(undecoded)) {
    cat(sprintf("   note: %d component(s) not decodable, scanned around: %s\n",
                length(undecoded), paste(undecoded, collapse = ", ")))
  }
  gaps <- 0L

  # --- 1. declared -----------------------------------------------------------
  tables <- dev_tables(wb)
  declared <- 0L
  missing_declared <- character(0)
  uncopied <- character(0)
  # A module named in a "form modules" row is imported, copied into the
  # document or form the row names, and then removed by Development, so it is
  # answered for by the copy and not by its own presence.
  form_sources <- character(0)
  for (tb in tables) {
    if (tb$tag == "general form modules") {
      form_sources <- c(form_sources, tolower(tb$first))
    }
  }
  for (tb in tables) {
    for (i in seq_along(tb$first)) {
      name <- tb$first[i]
      key <- tolower(name)
      declared <- declared + 1L
      if (tb$tag == "general form modules") {
        # The row names a module and the document or form its code is copied
        # into. The copy is what has to be there: every procedure the src
        # module declares, found in the target's own code.
        target <- tolower(tb$second[i])
        if (!nzchar(target)) {
          uncopied <- c(uncopied, sprintf("%s (no target named)", name))
        } else if (!(target %in% present)) {
          missing_declared <- c(missing_declared,
                                sprintf("%s (target of %s)", tb$second[i], name))
        } else if (!(key %in% names(candidate_files))) {
          uncopied <- c(uncopied, sprintf("%s -> %s (no src file to compare)",
                                          name, tb$second[i]))
        } else if (status_of[[target]] == "dumped") {
          wanted <- procedures_in(candidate_files[[key]])
          have <- procedures_in(file_of[[target]])
          gone <- setdiff(wanted, have)
          if (length(gone)) {
            shown <- paste(head(gone, 4L), collapse = ", ")
            if (length(gone) > 4L) {
              shown <- sprintf("%s and %d more", shown, length(gone) - 4L)
            }
            uncopied <- c(uncopied, sprintf(
              "%s -> %s (%d of %d procedures missing: %s)", name,
              tb$second[i], length(gone), length(wanted), shown))
          }
        }
        next
      }
      if (tb$tag == "general modules" && key %in% form_sources) next
      if (!(key %in% present)) {
        missing_declared <- c(missing_declared,
                              sprintf("%s (%s, folder %s)", name, tb$tag,
                                      tb$folder))
      }
      if (tb$tag == "general classes" && tolower(tb$second[i]) == "yes") {
        declared <- declared + 1L
        if (!(paste0("i", key) %in% present)) {
          missing_declared <- c(missing_declared,
                                sprintf("I%s (interface of %s, folder %s)",
                                        name, name, tb$folder))
        }
      }
    }
  }
  if (!length(tables)) {
    cat(sprintf("   %sno Dev tables%s: nothing declares what this workbook should carry\n",
                C_RED, C_OFF))
    gaps <- gaps + 1L
  } else {
    cat(sprintf("   declared: %d table(s), %d component(s)", length(tables),
                declared))
    if (attr(tables, "tests_seen") > 0L) {
      cat(sprintf(", %d tests table(s) not checked", attr(tables, "tests_seen")))
    }
    cat("\n")
    if (!length(missing_declared) && !length(uncopied)) {
      cat(sprintf("   %sdeclared%s: every component the Dev tables declare is in the workbook\n",
                  C_GREEN, C_OFF))
    } else {
      gaps <- gaps + 1L
      for (item in sort(missing_declared)) {
        cat(sprintf("   %sNOT IMPORTED %s%s\n", C_RED, item, C_OFF))
      }
      for (item in sort(uncopied)) {
        cat(sprintf("   %sNOT COPIED %s%s\n", C_RED, item, C_OFF))
      }
    }
  }

  # --- 2. closure ------------------------------------------------------------
  missing <- list()
  for (k in seq_along(comp_name)) {
    if (comp_status[k] != "dumped") next
    for (ref in references_in(comp_file[k], comp_name[k])) {
      if (ref %in% present) next
      missing[[ref]] <- c(missing[[ref]], comp_name[k])
    }
  }
  if (!length(missing)) {
    cat(sprintf("   %sclosed%s: every component the code names is in the workbook\n",
                C_GREEN, C_OFF))
  } else {
    gaps <- gaps + 1L
    for (ref in sort(names(missing))) {
      cat(sprintf("   %sMISSING %s%s  (%s)  named by: %s\n",
                  C_RED, candidate_names[[ref]], C_OFF, candidates[[ref]],
                  paste(sort(unique(missing[[ref]])), collapse = ", ")))
    }
  }

  # --- 3. callbacks ----------------------------------------------------------
  callbacks <- ribbon_callbacks(wb, work)
  if (is.null(callbacks)) {
    cat("   no ribbon: callbacks not checked\n")
  } else {
    module_procs <- character(0)
    for (k in seq_along(comp_name)) {
      if (comp_kind[k] != "Module" || comp_status[k] != "dumped") next
      module_procs <- c(module_procs, procedures_in(comp_file[k]))
    }
    unresolved <- callbacks[!(tolower(callbacks) %in% module_procs)]
    if (!length(unresolved)) {
      cat(sprintf("   %scallbacks%s: every ribbon callback (%d) has a procedure in a standard module\n",
                  C_GREEN, C_OFF, length(callbacks)))
    } else {
      gaps <- gaps + 1L
      cat(sprintf("   %sNO CALLBACK%s for %d of %d ribbon entries: %s\n",
                  C_RED, C_OFF, length(unresolved), length(callbacks),
                  paste(unresolved, collapse = ", ")))
    }
  }

  if (gaps > 0L) failed <- failed + 1L
  unlink(work, recursive = TRUE)
}

tone <- if (failed > 0L) C_RED else C_GREEN
cat(sprintf("\n%sdone: %d workbook(s) with a gap%s\n", tone, failed, C_OFF))
quit(status = failed)

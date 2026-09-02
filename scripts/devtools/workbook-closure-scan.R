#!/usr/bin/env Rscript
#
# workbook-closure-scan.R -- do the dev binaries carry every component their
# own code names?
#
#   Rscript scripts/devtools/workbook-closure-scan.R [workbook.xlsb ...]
#
# With no argument it scans the three dev binaries:
#   src/bin/setup/setup_dev.xlsb
#   src/bin/msetup/msetup_dev.xlsb
#   src/bin/designer/designer_dev.xlsb
#
# The setup, master setup and designer workbooks compile from whatever somebody
# last pasted into them, and a component left out is invisible until the first
# line that needs it dies with "Sub or Function not defined" -- a project-wide
# compile failure. Nothing else checks this: the test harness imports its own
# closure from src on every run, so it never meets the pasted set.
#
# The scan reads each workbook's ACTUAL code -- the pasted code is the ground
# truth, not the src tree -- through vba-inspect.R's dump mode, which
# decompresses every module stream of xl/vbaProject.bin in plain R, document
# modules included (ThisWorkbook carries the pasted event logic). No Excel, no
# COM, so it runs the same on Windows and macOS. String literals (doubled
# quotes understood) and comments are stripped, the rest is tokenised, and
# every reference to a known src component the workbook does not carry is
# reported with its referrers.
#
# The candidate universe is the file basenames of src/classes/** and
# src/modules/** minus stale/, so a name that lives inside another component --
# ProjectError inside Checking, an Office type like IRibbonControl -- is never
# a false alarm. This is a NAME scan, not a compiler: a local variable spelled
# exactly like a missing component can flag it, and a call built through
# Application.Run cannot be seen. Read the referrer list before acting.
#
# Exit code: the number of workbooks with at least one missing component.
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
cat(sprintf("candidate components in src: %d\n", length(candidates)))

# ---------------------------------------------------------------------------
# References inside one dumped module.
#
# A quote inside a VBA string is doubled, so the literal pattern eats whole
# strings before the apostrophe check -- "LLGeo" inside a TypeName compare
# never reads as a reference. VBA identifiers are case-insensitive, so the
# lookup is on the lowered token.
# ---------------------------------------------------------------------------
references_in <- function(path, self_name) {
  lines <- readLines(path, warn = FALSE, encoding = "latin1")
  lines <- gsub('"([^"]|"")*"', '""', lines)
  lines <- sub("'.*$", "", lines)
  tokens <- unlist(regmatches(lines,
                              gregexpr("[A-Za-z_][A-Za-z0-9_]*", lines)))
  tokens <- unique(tolower(tokens))
  tokens <- tokens[tokens %in% names(candidates)]
  setdiff(tokens, tolower(self_name))
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

  comp_name <- vapply(parsed, `[`, character(1), 2L)
  comp_file <- vapply(parsed, `[`, character(1), 3L)
  comp_status <- vapply(parsed, `[`, character(1), 4L)
  present <- tolower(comp_name)

  missing <- list()
  for (k in seq_along(comp_name)) {
    if (comp_status[k] != "dumped") next
    for (ref in references_in(comp_file[k], comp_name[k])) {
      if (ref %in% present) next
      missing[[ref]] <- c(missing[[ref]], comp_name[k])
    }
  }

  cat(sprintf("== %s : %d components\n", wb, length(comp_name)))
  undecoded <- comp_name[comp_status == "noextract"]
  if (length(undecoded)) {
    cat(sprintf("   note: %d component(s) not decodable, scanned around: %s\n",
                length(undecoded), paste(undecoded, collapse = ", ")))
  }

  if (!length(missing)) {
    cat(sprintf("   %sclosed%s: every component the code names is in the workbook\n",
                C_GREEN, C_OFF))
  } else {
    failed <- failed + 1L
    for (ref in sort(names(missing))) {
      cat(sprintf("   %sMISSING %s%s  (%s)  named by: %s\n",
                  C_RED, candidate_names[[ref]], C_OFF, candidates[[ref]],
                  paste(sort(unique(missing[[ref]])), collapse = ", ")))
    }
  }

  unlink(work, recursive = TRUE)
}

tone <- if (failed > 0L) C_RED else C_GREEN
cat(sprintf("\n%sdone: %d workbook(s) with a gap%s\n", tone, failed, C_OFF))
quit(status = failed)

#!/usr/bin/env Rscript
#
# linelist-components-scan.R -- every component the linelist needs is carried.
#
# Linelist.TransferAllCode moves a fixed list of classes, standard modules and
# forms out of the designer workbook into the linelist the user receives. A
# component left off that list is simply ABSENT from the delivered file, and
# VBA compiles a project as a whole: one unresolved type name costs the linelist
# its compile, so nothing in it works at all.
#
# transferred-code-scan.R is the other half of this check. It reads CALLS -- a
# transferred standard module invoking a Public procedure that lives in a module
# nobody transferred. This one reads NAMES: the class, module and enum
# identifiers the transferred set mentions, wherever they are mentioned from.
# Those are the two ways the delivered project can fail to resolve, and the
# call scan is blind to the second one. SectionMap, GeoFormCache and
# AnalysisRanges were all missing from the transfer list while that scan
# reported the set closed, because none of the three is reached by a bare call:
#
#   Dim secMap As SectionMap          a type in a declaration
#   Set secMap = SectionMap.Create(sh)  a predeclared class used statically
#   GeoFormCache.LoadFrom ThisWorkbook  the same, with no Dim anywhere
#
# WHAT IT READS
# -----------------------------------------------------------------------------
# The three Push blocks of TransferAllCode, straight off Linelist.cls, so the
# lists are never repeated here. Each name resolves to a file under src/classes
# or src/modules; a name that resolves to neither is the worksheet or workbook
# module the transfer passes by name, and is skipped.
#
# THE FORMS CARRY FormLogic CODE, NOT THE MODULES
# -----------------------------------------------------------------------------
# The ten .frm files are built by scripts/headless/merge-form-code.R, which
# writes the body of src/modules/linelistform/FormLogic<X>.bas into the code
# module of its form. So the code behind a transferred form is the FormLogic
# module that belongs to it, and this scan reads those .bas files as the source
# of the forms. That is why F_Geo needs GeoFormCache even though no transferred
# standard module mentions it.
#
# WHAT IT REPORTS
# -----------------------------------------------------------------------------
# Every identifier mentioned by the transferred set that (a) resolves to a
# class, module, Public Enum or Public Type defined somewhere in src/ and
# (b) is not itself in the transferred set. Names it cannot resolve inside this
# repository are VBA and Excel built-ins and cost no false alarm.
#
# Usage:
#   Rscript scripts/devtools/linelist-components-scan.R
#   Rscript scripts/devtools/linelist-components-scan.R --quiet   # findings only
#
# Exit code 1 when a finding is reported, 0 when the transferred set is closed.

args <- commandArgs(trailingOnly = TRUE)
quiet <- "--quiet" %in% args

repo_root <- normalizePath(file.path(dirname(sub("^--file=", "", grep(
  "^--file=", commandArgs(trailingOnly = FALSE), value = TRUE
)[1])), "..", ".."))
setwd(repo_root)

# ---------------------------------------------------------------------------
# Reading VBA source
# ---------------------------------------------------------------------------

# A quote inside a VBA string is doubled, so one pass that flips a flag on every
# quote gets the boundaries right. Everything after an unquoted apostrophe is a
# comment, and a comment naming a class is not a reference to it.
strip_comment <- function(line) {
  chars <- strsplit(line, "", fixed = TRUE)[[1]]
  if (!length(chars)) {
    return(line)
  }
  in_string <- FALSE
  for (idx in seq_along(chars)) {
    ch <- chars[idx]
    if (ch == "\"") {
      in_string <- !in_string
    } else if (ch == "'" && !in_string) {
      if (idx == 1) {
        return("")
      }
      return(paste(chars[seq_len(idx - 1)], collapse = ""))
    }
  }
  line
}

# A trailing underscore joins a statement to the next line. The joined text is
# what the compiler sees, so `Dim x As _ <newline> SectionMap` is one line here.
logical_lines <- function(path) {
  raw <- readLines(path, warn = FALSE)
  code <- vapply(raw, strip_comment, character(1), USE.NAMES = FALSE)
  joined <- character(0)
  numbers <- integer(0)
  buffer <- ""
  start <- 0L
  for (idx in seq_along(code)) {
    text <- code[idx]
    if (!nzchar(buffer)) {
      start <- idx
    }
    if (grepl("_[ \t]*$", text)) {
      buffer <- paste0(buffer, sub("_[ \t]*$", " ", text))
      next
    }
    joined <- c(joined, paste0(buffer, text))
    numbers <- c(numbers, start)
    buffer <- ""
  }
  if (nzchar(buffer)) {
    joined <- c(joined, buffer)
    numbers <- c(numbers, start)
  }
  list(text = joined, line = numbers)
}

src_files <- function(folder, pattern) {
  hits <- list.files(folder, pattern = pattern, recursive = TRUE,
                     full.names = TRUE)
  hits[!grepl("/stale/", hits, fixed = TRUE)]
}

ALL_CLASSES <- src_files("src/classes", "\\.cls$")
ALL_MODULES <- src_files("src/modules", "\\.bas$")

component_path <- function(name) {
  hits <- c(
    ALL_CLASSES[basename(ALL_CLASSES) == paste0(name, ".cls")],
    ALL_MODULES[basename(ALL_MODULES) == paste0(name, ".bas")]
  )
  if (!length(hits)) {
    return(NA_character_)
  }
  hits[1]
}

# ---------------------------------------------------------------------------
# The transferred set, read out of Linelist.cls
# ---------------------------------------------------------------------------

read_transfer_lists <- function() {
  lines <- logical_lines("src/classes/linelist/Linelist.cls")$text
  first <- grep("^\\s*Private Sub TransferAllCode", lines)
  if (!length(first)) {
    stop("TransferAllCode was not found in src/classes/linelist/Linelist.cls")
  }
  last <- grep("^\\s*End Sub", lines)
  last <- last[last > first[1]][1]
  body <- lines[first[1]:last]

  push_blocks <- body[grepl(
    "codesList\\.Push|codeTransfer\\.Transfer|codeTransfer\\.CopyModuleText", body
  )]
  unique(gsub("\"", "", unlist(lapply(push_blocks, function(line) {
    regmatches(line, gregexpr("\"[^\"]+\"", line))[[1]]
  }))))
}

# ---------------------------------------------------------------------------
# What each file DEFINES: its own name, plus its Public Enums and Types
# ---------------------------------------------------------------------------

# An Enum or Type declared in a class is a project-wide name, and the compile
# needs the file that declares it. That is its own bug class: an enum whose home
# class was left off the transfer is as fatal as a missing class.
public_type_names <- function(path) {
  lines <- logical_lines(path)$text
  hits <- regmatches(lines, regexec(
    "^\\s*(Public\\s+)?(Enum|Type)\\s+([A-Za-z_][A-Za-z0-9_]*)", lines
  ))
  names <- vapply(hits, function(m) if (length(m) >= 4) m[4] else NA_character_,
                  character(1))
  private <- grepl("^\\s*Private\\s+(Enum|Type)\\s", lines)
  unique(names[!is.na(names) & !private])
}

# name -> the file that defines it, for every class, module, Enum and Type in
# src/. Component names are registered last so a component always wins a clash
# with an enum of the same name.
definition_owner <- function() {
  owner <- new.env(parent = emptyenv())
  for (path in c(ALL_CLASSES, ALL_MODULES)) {
    for (nm in public_type_names(path)) {
      if (is.null(owner[[nm]])) assign(nm, path, envir = owner)
    }
  }
  for (path in c(ALL_CLASSES, ALL_MODULES)) {
    assign(sub("\\.(cls|bas)$", "", basename(path)), path, envir = owner)
  }
  owner
}

# ---------------------------------------------------------------------------
# What one file MENTIONS
# ---------------------------------------------------------------------------

# The three shapes that name a component without calling one of its procedures
# by a bare name, which is all transferred-code-scan.R can see:
#
#   As <Name>          a declaration, a parameter or a return type
#   New <Name>         an instantiation
#   <Name>.<member>    a qualified call, which is how every predeclared class
#                      factory and every cross-module call is written
DECL_PATTERN <- "\\bAs\\s+(?:New\\s+)?([A-Za-z_][A-Za-z0-9_]*)"
NEW_PATTERN <- "\\bNew\\s+([A-Za-z_][A-Za-z0-9_]*)"
QUALIFIED_PATTERN <- "(^|[^A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_]*)\\s*\\."

capture_all <- function(text, pattern, group) {
  hits <- gregexpr(pattern, text, perl = TRUE)
  matched <- regmatches(text, hits)[[1]]
  if (!length(matched)) {
    return(character(0))
  }
  vapply(matched, function(one) {
    sub(pattern, paste0("\\", group), one, perl = TRUE)
  }, character(1), USE.NAMES = FALSE)
}

mentions <- function(path) {
  parsed <- logical_lines(path)
  found <- list()
  for (idx in seq_along(parsed$text)) {
    text <- parsed$text[idx]
    if (!nzchar(trimws(text))) next
    if (grepl("^\\s*Attribute\\s", text)) next

    # A word inside a message is not a reference.
    text <- gsub("\"[^\"]*\"", "\"\"", text)

    names <- unique(c(
      capture_all(text, DECL_PATTERN, 1),
      capture_all(text, NEW_PATTERN, 1),
      capture_all(text, QUALIFIED_PATTERN, 2)
    ))
    for (nm in names) {
      found[[length(found) + 1]] <- list(name = nm, line = parsed$line[idx])
    }
  }
  found
}

# ---------------------------------------------------------------------------
# The files the transferred set is actually made of
# ---------------------------------------------------------------------------

# The form each FormLogic module is merged into, the table
# scripts/headless/merge-form-code.R writes from. Two of the pairs cannot be
# derived from the file name, so the table is explicit in both places.
FORM_LOGIC <- c(
  F_Advanced      = "FormLogicAdvanced",
  F_EpiWeek       = "FormLogicEpiWeek",
  F_Export        = "FormLogicExport",
  F_ExportMig     = "FormLogicExportMig",
  F_Geo           = "FormLogicGeo",
  F_ImportRep     = "FormLogicImportRep",
  F_ShowHideLL    = "FormLogicShowHide",
  F_ShowHidePrint = "FormLogicShowHidePrint",
  F_ShowHideSave  = "FormLogicShowHideSave",
  F_ShowVarLabels = "FormLogicShowVarLabels"
)

# The worksheet and workbook code modules the transfer injects by name. They are
# read too: EventLinelistSheet goes into every hlist2D worksheet of the
# linelist, so whatever it names has to travel as well.
INJECTED_MODULES <- c("EventLinelistWorkbook", "EventLinelistSheet")

transferred <- read_transfer_lists()

# name -> the file whose text the linelist will carry under that name.
carried <- list()
for (nm in transferred) {
  path <- component_path(nm)
  if (!is.na(path)) {
    carried[[nm]] <- path
    next
  }
  if (nm %in% names(FORM_LOGIC)) {
    logic <- component_path(FORM_LOGIC[[nm]])
    if (!is.na(logic)) carried[[nm]] <- logic
  }
}
for (nm in INJECTED_MODULES) {
  path <- component_path(nm)
  if (!is.na(path)) carried[[nm]] <- path
}

# Everything the delivered project will be able to resolve: the components it
# carries, plus every Enum and Type those components declare.
resolvable <- unique(c(
  names(carried),
  unlist(lapply(carried, public_type_names), use.names = FALSE)
))

owner <- definition_owner()

findings <- list()
for (nm in names(carried)) {
  path <- carried[[nm]]
  for (hit in mentions(path)) {
    if (hit$name %in% resolvable) next
    home <- owner[[hit$name]]
    if (is.null(home)) next
    findings[[length(findings) + 1]] <- data.frame(
      carrier = nm,
      source = sub("^src/", "", path),
      line = hit$line,
      needs = hit$name,
      defined_in = sub("^src/", "", home),
      stringsAsFactors = FALSE
    )
  }
}

if (!quiet) {
  cat("Components in the transfer list:", length(transferred), "\n")
  cat("Source files behind them:", length(carried), "\n")
  cat("Names the delivered project can resolve:", length(resolvable), "\n\n")
}

if (!length(findings)) {
  cat("The transferred set is closed: every name it mentions travels with it.\n")
  quit(status = 0)
}

report <- do.call(rbind, findings)
report <- report[order(report$needs, report$carrier, report$line), ]

# One line per missing component, with every carrier that asks for it: the fix
# is one name added to TransferAllCode, not one per call site.
cat("The transferred set mentions a component the generated linelist will not carry.\n")
cat("VBA compiles a project as a whole, so the linelist fails to compile.\n\n")
for (needed in unique(report$needs)) {
  rows <- report[report$needs == needed, ]
  cat(sprintf("  %s  (defined in %s)\n", needed, rows$defined_in[1]))
  for (idx in seq_len(nrow(rows))) {
    cat(sprintf("      wanted by %s at %s:%d\n",
                rows$carrier[idx], rows$source[idx], rows$line[idx]))
  }
}
cat("\nAdd each one to the matching list in Linelist.TransferAllCode.\n")

quit(status = 1)

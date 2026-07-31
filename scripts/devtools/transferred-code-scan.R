#!/usr/bin/env Rscript
#
# transferred-code-scan.R -- check that a generated linelist can compile.
#
# Linelist.TransferAllCode moves a fixed list of classes, standard modules and
# forms out of the designer workbook and into the linelist the user receives.
# Anything left off that list is simply absent from the delivered file, and the
# first line of VBA that runs in it dies with "Sub or Function not defined" --
# a project-wide compile failure, so nothing in the linelist works at all.
#
# Nothing catches that. The designer compiles because every module is there; the
# test harness never builds a linelist. This scan is the free half of that
# check: it reads the three lists out of Linelist.cls, works out which Public
# procedures the transferred set defines, and reports every call in a
# transferred standard module that resolves to a Public procedure defined
# somewhere in src/ that is NOT transferred.
#
# It reports only calls whose callee it can see in this repo, so Excel and VBA
# built-in procedures cost no false alarms. The nine .frm files live inside the
# designer workbook and are not in git, so the code behind the forms is out of
# reach here and stays a check by eye.
#
# Usage:
#   Rscript scripts/devtools/transferred-code-scan.R
#   Rscript scripts/devtools/transferred-code-scan.R --quiet   # findings only
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
# comment.
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
# what the compiler sees, so calls are matched against it.
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

component_path <- function(name) {
  hits <- c(
    list.files("src/classes", pattern = paste0("^", name, "\\.cls$"),
               recursive = TRUE, full.names = TRUE),
    list.files("src/modules", pattern = paste0("^", name, "\\.bas$"),
               recursive = TRUE, full.names = TRUE)
  )
  hits <- hits[!grepl("/stale/", hits, fixed = TRUE)]
  if (!length(hits)) {
    return(NA_character_)
  }
  hits[1]
}

# ---------------------------------------------------------------------------
# The transferred set, read out of Linelist.cls
# ---------------------------------------------------------------------------

# Each Push block in TransferAllCode lists one kind of component. The names are
# string literals, so they are read straight off the source and the lists never
# have to be repeated here.
read_transfer_lists <- function() {
  lines <- logical_lines("src/classes/linelist/Linelist.cls")$text
  first <- grep("^\\s*Private Sub TransferAllCode", lines)
  if (!length(first)) {
    stop("TransferAllCode was not found in src/classes/linelist/Linelist.cls")
  }
  last <- grep("^\\s*End Sub", lines)
  last <- last[last > first[1]][1]
  body <- lines[first[1]:last]

  literals <- function(pattern) {
    hits <- body[grepl(pattern, body)]
    unlist(lapply(hits, function(line) {
      regmatches(line, gregexpr("\"[^\"]+\"", line))[[1]]
    }))
  }

  clean <- function(values) unique(gsub("\"", "", values))

  push_blocks <- body[grepl("codesList\\.Push|codeTransfer\\.Transfer|codeTransfer\\.CopyModuleText", body)]
  all_names <- clean(unlist(lapply(push_blocks, function(line) {
    regmatches(line, gregexpr("\"[^\"]+\"", line))[[1]]
  })))

  # A name that resolves to a file in src/ is a component; the rest are the
  # worksheet and workbook module names the transfer passes by name.
  paths <- vapply(all_names, component_path, character(1), USE.NAMES = TRUE)
  list(names = all_names, paths = paths)
}

# ---------------------------------------------------------------------------
# Public procedures, per component
# ---------------------------------------------------------------------------

PROC_PATTERN <- paste0(
  "^\\s*(Public\\s+|Friend\\s+)?(Static\\s+)?",
  "(Sub|Function|Property\\s+(Get|Let|Set))\\s+([A-Za-z_][A-Za-z0-9_]*)"
)

public_procs <- function(path) {
  if (is.na(path) || !file.exists(path)) {
    return(character(0))
  }
  lines <- logical_lines(path)$text
  if (any(grepl("^\\s*Option\\s+Private\\s+Module", lines))) {
    return(character(0))
  }
  hits <- regmatches(lines, regexec(PROC_PATTERN, lines))
  names <- vapply(hits, function(m) if (length(m) >= 6) m[6] else NA_character_,
                  character(1))
  names <- names[!is.na(names)]
  # A procedure with no modifier is Public in VBA, and Private is written out.
  private <- grepl("^\\s*Private\\s+", lines)
  keep <- vapply(seq_along(lines), function(idx) {
    grepl(PROC_PATTERN, lines[idx]) && !private[idx]
  }, logical(1))
  unique(vapply(hits[keep], function(m) m[6], character(1)))
}

# ---------------------------------------------------------------------------
# Calls made by the transferred standard modules
# ---------------------------------------------------------------------------

# A bare call is an identifier at the head of a statement, or one that follows
# Call / Then / Else / =. A member call carries a dot in front of it and belongs
# to an object, so it is left out here.
KEYWORDS <- c(
  "If", "Then", "Else", "ElseIf", "End", "Exit", "For", "Next", "Do", "Loop",
  "While", "Wend", "Select", "Case", "With", "Set", "Let", "Dim", "Const",
  "Public", "Private", "Friend", "Static", "Sub", "Function", "Property",
  "Get", "On", "Error", "Resume", "GoTo", "Call", "Each", "In", "To", "Step",
  "Option", "Explicit", "Type", "Enum", "Declare", "ReDim", "Erase", "Stop",
  "Return", "Attribute", "New", "Nothing", "True", "False", "Not", "And",
  "Or", "Is", "Like", "Mod", "ByVal", "ByRef", "Optional", "ParamArray",
  "As", "Preserve", "Me", "Application", "Debug", "Err", "Excel"
)

calls_made <- function(path) {
  parsed <- logical_lines(path)
  found <- list()
  for (idx in seq_along(parsed$text)) {
    text <- parsed$text[idx]
    if (!nzchar(trimws(text))) next
    if (grepl(PROC_PATTERN, text)) next

    # Drop string literals so a word inside a message is never read as a call.
    text <- gsub("\"[^\"]*\"", "\"\"", text)

    # Anything reached through a dot belongs to an object.
    text <- gsub("[A-Za-z0-9_]+\\s*\\.\\s*[A-Za-z_][A-Za-z0-9_]*", " ", text)

    tokens <- regmatches(text, gregexpr("[A-Za-z_][A-Za-z0-9_]*", text))[[1]]
    tokens <- setdiff(tokens, KEYWORDS)
    for (token in tokens) {
      found[[length(found) + 1]] <- list(name = token, line = parsed$line[idx])
    }
  }
  found
}

# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------

transfer <- read_transfer_lists()
transferred_paths <- transfer$paths[!is.na(transfer$paths)]

transferred_procs <- unique(unlist(lapply(transferred_paths, public_procs)))

# Every Public procedure defined anywhere in src/, and where it lives.
all_modules <- list.files("src/modules", pattern = "\\.bas$",
                          recursive = TRUE, full.names = TRUE)
all_modules <- all_modules[!grepl("/stale/", all_modules, fixed = TRUE)]

owner <- new.env(parent = emptyenv())
for (path in all_modules) {
  for (proc in public_procs(path)) {
    if (is.null(owner[[proc]])) {
      assign(proc, path, envir = owner)
    }
  }
}

# The callers: the transferred standard modules, plus the ribbon module, which
# is transferred on one branch and copied on the other.
caller_paths <- transferred_paths[grepl("\\.bas$", transferred_paths)]
ribbon <- component_path("EventsLinelistRibbon")
if (!is.na(ribbon)) caller_paths <- unique(c(caller_paths, ribbon))

findings <- list()
for (path in caller_paths) {
  for (call in calls_made(path)) {
    if (call$name %in% transferred_procs) next
    home <- owner[[call$name]]
    if (is.null(home)) next
    findings[[length(findings) + 1]] <- data.frame(
      caller = sub("^src/", "", path),
      line = call$line,
      callee = call$name,
      defined_in = sub("^src/", "", home),
      stringsAsFactors = FALSE
    )
  }
}

if (!quiet) {
  cat("Transferred components:", length(transferred_paths), "\n")
  cat("Public procedures in the transferred set:", length(transferred_procs), "\n")
  cat("Standard modules scanned for calls:", length(caller_paths), "\n\n")
}

if (!length(findings)) {
  cat("The transferred set is closed: every call resolves inside it.\n")
  quit(status = 0)
}

report <- unique(do.call(rbind, findings))
report <- report[order(report$caller, report$line), ]

cat("A transferred module calls a procedure the generated linelist will not carry.\n")
cat("The linelist fails project compilation on the first line of VBA that runs.\n\n")
for (idx in seq_len(nrow(report))) {
  row <- report[idx, ]
  cat(sprintf("  %s:%d  calls %s, defined in %s\n",
              row$caller, row$line, row$callee, row$defined_in))
}
cat("\nAdd the module holding the callee to the module list in ",
    "Linelist.TransferAllCode.\n", sep = "")

quit(status = 1)

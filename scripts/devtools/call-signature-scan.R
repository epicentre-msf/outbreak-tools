#!/usr/bin/env Rscript
#
# call-signature-scan.R -- check every early-bound call against its callee.
#
# VBA catches two faults at COMPILE time, and only when the object is typed:
#
#   1. a call that leaves out a required argument ("Argument not optional"),
#   2. a call to a member the class does not have ("Method or data member not
#      found").
#
# A class with no registry row is never imported into the test workbook, so
# nothing compiles it and both faults sit in the tree for weeks. This scan is
# the free half of that compiler: it reads the sources, works out the type of
# every object variable it can, and checks each `<obj>.<Member>` call against
# the real signature.
#
# It reports only calls whose receiver resolves to a class in this repo, so
# Excel and VBA built-in members are out of reach and cost no false alarms.
#
# Usage:
#   Rscript scripts/devtools/call-signature-scan.R            # classes + modules
#   Rscript scripts/devtools/call-signature-scan.R --tests    # also src/tests
#   Rscript scripts/devtools/call-signature-scan.R --quiet    # findings only
#
# Exit code 1 when a finding is reported, 0 when the tree is clean.

args <- commandArgs(trailingOnly = TRUE)
scan_tests <- "--tests" %in% args
quiet <- "--quiet" %in% args

repo_root <- normalizePath(file.path(dirname(sub("^--file=", "", grep(
  "^--file=", commandArgs(trailingOnly = FALSE), value = TRUE
)[1])), "..", ".."))
setwd(repo_root)

# ---------------------------------------------------------------------------
# Reading a VBA source file into logical lines
# ---------------------------------------------------------------------------

# A quote inside a VBA string is doubled, so a single pass that flips a flag on
# every quote gets the string boundaries right. Everything after an unquoted
# apostrophe is a comment.
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

# Joins continuation lines and remembers the line number each logical line
# started on, so a finding points at the line a reader can open.
logical_lines <- function(path) {
  raw <- readLines(path, warn = FALSE)
  raw <- gsub("\r$", "", raw)

  texts <- character(0)
  starts <- integer(0)
  buffer <- ""
  start_at <- 0L

  for (idx in seq_along(raw)) {
    line <- strip_comment(raw[idx])
    if (buffer == "") start_at <- idx
    continued <- grepl("[ \t]_[ \t]*$", line)
    if (continued) line <- sub("[ \t]_[ \t]*$", " ", line)
    buffer <- paste0(buffer, line)
    if (!continued) {
      texts <- c(texts, trimws(buffer))
      starts <- c(starts, start_at)
      buffer <- ""
    }
  }
  if (buffer != "") {
    texts <- c(texts, trimws(buffer))
    starts <- c(starts, start_at)
  }

  # A colon joins several statements on one line. Splitting on it would break
  # a `Dim x As String: x = "a:b"`, so only the simple case is split.
  list(text = texts, line = starts)
}

# ---------------------------------------------------------------------------
# Signatures
# ---------------------------------------------------------------------------

# Splits an argument list at commas that sit outside brackets and strings.
split_top_level <- function(text) {
  chars <- strsplit(text, "", fixed = TRUE)[[1]]
  if (!length(chars)) {
    return(character(0))
  }
  depth <- 0L
  in_string <- FALSE
  parts <- character(0)
  current <- ""
  for (ch in chars) {
    if (ch == "\"") in_string <- !in_string
    if (!in_string) {
      if (ch == "(" || ch == "[") depth <- depth + 1L
      if (ch == ")" || ch == "]") depth <- depth - 1L
      if (ch == "," && depth == 0L) {
        parts <- c(parts, current)
        current <- ""
        next
      }
    }
    current <- paste0(current, ch)
  }
  c(parts, current)
}

parse_params <- function(param_text) {
  param_text <- trimws(param_text)
  if (param_text == "") {
    return(list())
  }
  pieces <- split_top_level(param_text)
  out <- list()
  for (piece in pieces) {
    piece <- trimws(piece)
    if (piece == "") next
    optional <- grepl("^Optional\\b", piece, ignore.case = TRUE)
    variadic <- grepl("^ParamArray\\b", piece, ignore.case = TRUE)
    piece <- sub("^(Optional|ParamArray)\\s+", "", piece, ignore.case = TRUE)
    piece <- sub("^(ByVal|ByRef)\\s+", "", piece, ignore.case = TRUE)
    name <- sub("^([A-Za-z_][A-Za-z0-9_]*).*$", "\\1", piece)
    out[[length(out) + 1]] <- list(
      name = name,
      optional = optional || variadic,
      variadic = variadic
    )
  }
  out
}

proc_pattern <- paste0(
  "^(?:Public\\s+|Private\\s+|Friend\\s+|Global\\s+)?",
  "(?:Static\\s+)?",
  "(Sub|Function|Property\\s+Get|Property\\s+Let|Property\\s+Set)\\s+",
  "([A-Za-z_][A-Za-z0-9_]*)\\s*\\((.*)$"
)

# The signature of one component: every procedure it declares, with the
# parameters each one requires.
parse_component <- function(path) {
  lines <- logical_lines(path)
  procs <- list()

  for (idx in seq_along(lines$text)) {
    text <- lines$text[idx]
    if (!grepl(proc_pattern, text, perl = TRUE, ignore.case = TRUE)) next

    kind <- sub(proc_pattern, "\\1", text, perl = TRUE, ignore.case = TRUE)
    name <- sub(proc_pattern, "\\2", text, perl = TRUE, ignore.case = TRUE)
    tail <- sub(proc_pattern, "\\3", text, perl = TRUE, ignore.case = TRUE)

    # Everything up to the closing bracket of the parameter list.
    closing <- regexpr("\\)[^)]*$", tail)
    param_text <- if (closing > 0) substr(tail, 1, closing - 1) else tail

    kind <- gsub("\\s+", " ", tolower(kind))
    key <- tolower(name)
    params <- parse_params(param_text)
    required <- sum(!vapply(params, function(p) p$optional, logical(1)))

    if (is.null(procs[[key]])) {
      procs[[key]] <- list(
        name = name, kinds = character(0), params = params,
        required = required, line = lines$line[idx]
      )
    }
    procs[[key]]$kinds <- c(procs[[key]]$kinds, kind)
    # A Property Get and its Let carry different lists. The Get is what a read
    # resolves against, so it wins when both are declared.
    if (kind == "property get" || !("property get" %in% procs[[key]]$kinds)) {
      procs[[key]]$params <- params
      procs[[key]]$required <- required
    }
  }
  procs
}

# ---------------------------------------------------------------------------
# Types of the variables in one file
# ---------------------------------------------------------------------------

decl_pattern <- paste0(
  "^(?:Dim|Private|Public|Global|Static)\\s+",
  "([A-Za-z_][A-Za-z0-9_]*)\\s+As\\s+(?:New\\s+)?",
  "([A-Za-z_][A-Za-z0-9_]*)\\s*$"
)
field_pattern <- "^([A-Za-z_][A-Za-z0-9_]*)\\s+As\\s+(?:New\\s+)?([A-Za-z_][A-Za-z0-9_]*)\\s*$"

# The declarations of one line, folded into a scope map. A variable typed as a
# record contributes one entry per field the record holds, so `this.specs`
# resolves the same way a plain local does.
add_declaration <- function(scope, text, known, records) {
  if (!grepl(decl_pattern, text, perl = TRUE, ignore.case = TRUE)) {
    return(scope)
  }
  vname <- tolower(sub(decl_pattern, "\\1", text, perl = TRUE, ignore.case = TRUE))
  vtype <- sub(decl_pattern, "\\2", text, perl = TRUE, ignore.case = TRUE)

  if (vtype %in% known) {
    scope[[vname]] <- vtype
  } else if (vtype %in% names(records)) {
    for (fname in names(records[[vtype]])) {
      scope[[paste0(vname, ".", fname)]] <- records[[vtype]][[fname]]
    }
  } else {
    # A name reused for another type in a later procedure must stop resolving,
    # so the entry is dropped rather than left pointing at the old class.
    scope[[vname]] <- NULL
  }
  scope
}

# The record types a file declares, and the class type of each of their fields.
record_types <- function(lines, known) {
  records <- list()
  current <- NULL

  for (text in lines$text) {
    if (grepl("^(Private|Public)\\s+Type\\s+", text, ignore.case = TRUE)) {
      current <- sub(
        "^(?:Private|Public)\\s+Type\\s+([A-Za-z_][A-Za-z0-9_]*).*$", "\\1",
        text,
        ignore.case = TRUE
      )
      records[[current]] <- list()
      next
    }
    if (grepl("^End\\s+Type\\b", text, ignore.case = TRUE)) {
      current <- NULL
      next
    }
    if (!is.null(current) && grepl(field_pattern, text, perl = TRUE)) {
      fname <- sub(field_pattern, "\\1", text, perl = TRUE)
      ftype <- sub(field_pattern, "\\2", text, perl = TRUE)
      if (ftype %in% known) records[[current]][[tolower(fname)]] <- ftype
    }
  }
  records
}

# Module-level fields, which every procedure of the file can see. A declaration
# inside a procedure belongs to that procedure alone and is collected during
# the scan itself.
module_types <- function(lines, known, records) {
  scope <- list()
  in_proc <- FALSE
  in_record <- FALSE

  for (text in lines$text) {
    if (grepl("^(Private|Public)\\s+Type\\s+", text, ignore.case = TRUE)) in_record <- TRUE
    if (grepl("^End\\s+Type\\b", text, ignore.case = TRUE)) {
      in_record <- FALSE
      next
    }
    if (in_record) next

    if (grepl(proc_pattern, text, perl = TRUE, ignore.case = TRUE)) {
      in_proc <- TRUE
      next
    }
    if (grepl("^End\\s+(Sub|Function|Property)\\b", text, ignore.case = TRUE)) {
      in_proc <- FALSE
      next
    }
    if (!in_proc) scope <- add_declaration(scope, text, known, records)
  }
  scope
}

# ---------------------------------------------------------------------------
# Calls
# ---------------------------------------------------------------------------

# The text of the argument list of a call, and whether the call is a property
# assignment. A call written with brackets ends at the matching bracket; a
# statement call runs to the end of the logical line.
call_arguments <- function(text, after, statement_start) {
  rest <- substr(text, after, nchar(text))
  lead <- sub("^([ \t]*).*$", "\\1", rest)
  rest <- substr(rest, nchar(lead) + 1, nchar(rest))

  if (startsWith(rest, "=") && !startsWith(rest, "==")) {
    return(list(kind = "assign", args = character(0)))
  }
  if (startsWith(rest, "(")) {
    chars <- strsplit(rest, "", fixed = TRUE)[[1]]
    depth <- 0L
    in_string <- FALSE
    for (idx in seq_along(chars)) {
      ch <- chars[idx]
      if (ch == "\"") in_string <- !in_string
      if (!in_string) {
        if (ch == "(") depth <- depth + 1L
        if (ch == ")") {
          depth <- depth - 1L
          if (depth == 0L) {
            tail_text <- trimws(substr(rest, idx + 1, nchar(rest)))
            # `Assert.IsTrue (a Is b), "message"` is a statement call whose
            # first argument happens to be bracketed. In an expression the same
            # bracket is the whole argument list, and whatever follows belongs
            # to the expression around it.
            if (statement_start && tail_text != "") {
              if (startsWith(tail_text, "=") && !startsWith(tail_text, "==")) {
                return(list(kind = "assign", args = character(0)))
              }
              return(list(kind = "call", args = split_top_level(rest)))
            }
            inner <- substr(rest, 2, idx - 1)
            return(list(kind = "call", args = split_top_level(inner)))
          }
        }
      }
    }
    return(list(kind = "unknown", args = character(0)))
  }
  if (rest == "" || startsWith(rest, ".")) {
    return(list(kind = "bare", args = character(0)))
  }
  list(kind = "call", args = split_top_level(rest))
}

named_arguments <- function(args) {
  named <- character(0)
  positional <- 0L
  for (arg in args) {
    arg <- trimws(arg)
    if (arg == "") next
    if (grepl("^[A-Za-z_][A-Za-z0-9_]*\\s*:=", arg)) {
      named <- c(named, tolower(sub("^([A-Za-z_][A-Za-z0-9_]*)\\s*:=.*$", "\\1", arg)))
    } else {
      positional <- positional + 1L
    }
  }
  list(named = named, positional = positional)
}

# ---------------------------------------------------------------------------
# The scan
# ---------------------------------------------------------------------------

source_files <- function() {
  files <- c(
    list.files("src/classes", pattern = "\\.cls$", recursive = TRUE, full.names = TRUE),
    list.files("src/modules", pattern = "\\.bas$", recursive = TRUE, full.names = TRUE)
  )
  if (scan_tests) {
    files <- c(files, list.files("src/tests",
      pattern = "\\.(bas|cls)$",
      recursive = TRUE, full.names = TRUE
    ))
  }
  files[!grepl("/stale/", files, fixed = TRUE)]
}

files <- source_files()
components <- setNames(lapply(files, parse_component), tools::file_path_sans_ext(basename(files)))
known <- names(components)

findings <- list()
add_finding <- function(file, line, message) {
  findings[[length(findings) + 1]] <<- list(file = file, line = line, message = message)
}

member_pattern <- "(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_]*(?:\\.[A-Za-z_][A-Za-z0-9_]*)?)\\.([A-Za-z_][A-Za-z0-9_]*)"

# The text of a line with every string literal blanked out, keeping the length
# so the offsets of a match still index into the real line. A file name written
# as a literal reads exactly like a member call.
blank_strings <- function(text) {
  chars <- strsplit(text, "", fixed = TRUE)[[1]]
  if (!length(chars)) {
    return(text)
  }
  in_string <- FALSE
  for (idx in seq_along(chars)) {
    ch <- chars[idx]
    if (ch == "\"") {
      in_string <- !in_string
      next
    }
    if (in_string) chars[idx] <- " "
  }
  paste(chars, collapse = "")
}

for (path in files) {
  owner <- tools::file_path_sans_ext(basename(path))
  lines <- logical_lines(path)
  records <- record_types(lines, known)
  fields <- module_types(lines, known, records)
  locals <- list()

  for (idx in seq_along(lines$text)) {
    text <- lines$text[idx]
    if (text == "") next

    if (grepl(proc_pattern, text, perl = TRUE, ignore.case = TRUE)) {
      locals <- list()
    } else if (grepl("^End\\s+(Sub|Function|Property)\\b", text, ignore.case = TRUE)) {
      locals <- list()
    } else if (grepl("^Dim\\s+", text, ignore.case = TRUE)) {
      locals <- add_declaration(locals, text, known, records)
    }

    if (grepl("^(Dim|Private|Public|Global|Static|Const|Type|End|Declare)\\b", text, ignore.case = TRUE)) next

    matches <- gregexpr(member_pattern, blank_strings(text), perl = TRUE)[[1]]
    if (matches[1] == -1) next
    starts <- attr(matches, "capture.start")
    lengths <- attr(matches, "capture.length")

    for (row in seq_len(nrow(starts))) {
      receiver <- substr(text, starts[row, 1], starts[row, 1] + lengths[row, 1] - 1)
      member <- substr(text, starts[row, 2], starts[row, 2] + lengths[row, 2] - 1)
      after <- starts[row, 2] + lengths[row, 2]

      target <- NULL
      key <- tolower(receiver)
      if (receiver %in% known) {
        target <- receiver # a PredeclaredId factory call
      } else if (!is.null(locals[[key]])) {
        target <- locals[[key]]
      } else if (!is.null(fields[[key]])) {
        target <- fields[[key]]
      }
      if (is.null(target)) next
      if (target == owner) next # Me-style self calls reach private members too

      procs <- components[[target]]
      signature <- procs[[tolower(member)]]

      if (is.null(signature)) {
        add_finding(path, lines$line[idx], paste0(
          target, " has no member ", member, " (", receiver, ".", member, ")"
        ))
        next
      }

      lead_text <- trimws(substr(text, 1, starts[row, 1] - 1))
      statement_start <- (lead_text == "" || grepl("^Call$", lead_text, ignore.case = TRUE))

      call <- call_arguments(text, after, statement_start)
      if (call$kind != "call") next

      supplied <- named_arguments(call$args)
      if (any(vapply(signature$params, function(p) p$variadic, logical(1)))) next

      required <- Filter(function(p) !p$optional, signature$params)
      if (!length(required)) next

      if (length(supplied$named)) {
        missing <- Filter(function(p) !(tolower(p$name) %in% supplied$named), required)
        # A call may mix positional and named arguments; the positional ones
        # fill the list from the front.
        if (supplied$positional > 0) {
          missing <- if (supplied$positional >= length(missing)) list() else missing
        }
        if (length(missing)) {
          add_finding(path, lines$line[idx], paste0(
            target, ".", member, " needs ",
            paste(vapply(missing, function(p) p$name, character(1)), collapse = ", "),
            " and the call does not pass ",
            if (length(missing) > 1) "them" else "it"
          ))
        }
      } else if (supplied$positional < length(required)) {
        add_finding(path, lines$line[idx], paste0(
          target, ".", member, " takes ", length(required),
          " required arguments and the call passes ", supplied$positional
        ))
      }
    }
  }
}

if (!quiet) {
  cat("Scanned", length(files), "components for missing members and missing arguments.\n")
}

if (!length(findings)) {
  if (!quiet) cat("Clean.\n")
  quit(status = 0)
}

cat("\n", length(findings), " finding(s):\n\n", sep = "")
for (finding in findings) {
  cat(finding$file, ":", finding$line, "  ", finding$message, "\n", sep = "")
}
quit(status = 1)

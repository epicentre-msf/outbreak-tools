#!/usr/bin/env Rscript
#
# name-shadowing-scan.R -- a variable must not be named after a type, and it
# must not be named after a procedure of its own module either.
#
# TWO FAULTS, ONE CAUSE
# -----------------------------------------------------------------------------
# VBA is not case sensitive and it resolves an identifier to the nearest scope.
# So a local can take the name of something the module already has, and every
# later mention of that name means the local. The first half of this scan is
# about names of TYPES. The second is about names of PROCEDURES, which fails the
# same way and reads even more innocently:
#
#   Public Function Create(ByVal sh As Worksheet, _
#                          Optional ByVal checkRequirements As Boolean = True)
#       ...
#       If checkRequirements Then CheckRequirements sh, geoStore, levelStore
#
# CheckRequirements is a Private Sub of that class. The parameter has taken its
# name, so the second CheckRequirements on that line is the PARAMETER, and the
# statement becomes an index into a Boolean rather than a call. It compiles. At
# run time the checks never run, or the line raises somewhere that reads as
# unrelated. This one cost a wedged headless run: no results file at all,
# because the module died before its cleanup.
#
# VBA is not case sensitive. `Dim passwords As Passwords` therefore declares a
# variable whose name IS the class name, and from that point on, inside that
# scope, the identifier `Passwords` means the variable. The class has become
# unreachable by name:
#
#   Private showHideLayout As ShowHideLayout      'module scope
#   ...
#   Set LayoutFor = ShowHideLayout.Create(sh, ...)   '<- reads the VARIABLE
#
# That compiles. The variable is of the class's type, so it offers the same
# `Create` member, and the compiler is satisfied. It fails at RUN time with
# error 91, "Object variable or With block variable not set", pointing at a line
# that looks correct, because the variable is Nothing. Chasing that error leads
# everywhere except the declaration that caused it.
#
# WHAT COUNTS AS A TYPE
# -----------------------------------------------------------------------------
# Every identifier this repository ever writes after `As`. That set is built
# from the source itself rather than listed here, so it covers the project's own
# classes and the host's types alike -- Workbook, ListObject, FormatCondition --
# without a hand-kept list going stale. Public Enum and Public Type names are
# added from their declarations, and so are Enum MEMBERS, which live in the
# project namespace too and are shadowed the same way.
#
# HOW BAD EACH HIT IS
# -----------------------------------------------------------------------------
#   breaking  the shadowed TYPE is used as a qualifier (`Name.Member`) or after
#             `New` inside the scope the shadow covers. This one is live: the
#             read reaches the variable. Error 91 at run time.
#   calling   the shadowed PROCEDURE is called by name inside the scope the
#             shadow covers. Also live: the call reaches the variable instead of
#             the procedure, and the procedure silently never runs.
#   error     a shadow of a name this project defines, with no such read found.
#             Nothing is broken today and the next line added breaks it.
#   procname  a shadow of a procedure of the same module, with no such call
#             found in the scope. Same rule, waiting for the next line.
#   builtin   a shadow of a host type (Workbook, Range, ListObject). Same rule,
#             lower blast radius, and there are many.
#   member    a User Defined Type field. A field name is scoped to the type, so
#             `this.passwords` and `Passwords.Create` coexist. Banned anyway --
#             it reads as a shadow to everyone who meets it.
#
# WHY THE PROCEDURE CHECK IS PER MODULE
# -----------------------------------------------------------------------------
# A procedure of ANOTHER class is reached through an object, so it cannot be
# shadowed by a local here. Only the module's own procedures are in scope
# unqualified, so only those are collected, once per file.
#
# SCOPE, AND WHY IT MATTERS FOR THE VERDICT
# -----------------------------------------------------------------------------
# A declaration above the first procedure covers the whole module. One inside a
# procedure covers that procedure. The scan tracks which, because it decides
# where to look for the read that makes a shadow breaking.
#
# Usage:
#   Rscript scripts/devtools/name-shadowing-scan.R
#   Rscript scripts/devtools/name-shadowing-scan.R --quiet      # findings only
#   Rscript scripts/devtools/name-shadowing-scan.R --breaking   # live findings only
#                                                               # (breaking + calling)
#   Rscript scripts/devtools/name-shadowing-scan.R --path src/modules/linelist
#   Rscript scripts/devtools/name-shadowing-scan.R --strict     # procname gates too
#
# Exit code 1 when a finding is reported at the requested severity or above.
#
# `procname` is PRINTED but does not set the exit code unless --strict is given.
# There are dozens of them and none is broken today, and this file already says
# why that matters: a scan that always reports findings is a scan people stop
# reading. The gate stays on the classes that were gating before, plus `calling`
# -- the live one, the one that costs a run.

args <- commandArgs(trailingOnly = TRUE)
quiet <- "--quiet" %in% args
only_breaking <- "--breaking" %in% args
strict <- "--strict" %in% args

repo_root <- normalizePath(file.path(dirname(sub("^--file=", "", grep(
  "^--file=", commandArgs(trailingOnly = FALSE), value = TRUE
)[1])), "..", ".."))
setwd(repo_root)

path_arg <- NULL
if ("--path" %in% args) {
  i <- which(args == "--path")
  if (length(args) > i) path_arg <- args[i + 1]
}

say <- function(...) if (!quiet) cat(..., "\n", sep = "")

# ---------------------------------------------------------------------------
# Reading VBA source
# ---------------------------------------------------------------------------

# Strip the trailing comment, leaving string literals alone. A ' inside quotes
# is a character, not a comment, and `Value = "won't"` must not lose its tail.
strip_comment <- function(line) {
  chars <- strsplit(line, "", fixed = TRUE)[[1]]
  in_string <- FALSE
  for (i in seq_along(chars)) {
    if (chars[i] == '"') {
      in_string <- !in_string
    } else if (chars[i] == "'" && !in_string) {
      return(if (i == 1) "" else paste(chars[seq_len(i - 1)], collapse = ""))
    }
  }
  line
}

# Blank the inside of every string literal. This project writes its messages and
# its test names in English, and English says "as expected", "reported as
# existing", "read as row 1". Every one of those reads as a type declaration to
# a regex, and each one would put a common word into the forbidden set and then
# condemn every variable that used it. Assertions alone contributed `written`,
# `missing`, `skipped`, `changed` and `expected`.
strip_strings <- function(line) {
  gsub('"[^"]*"', '""', line, perl = TRUE)
}

# The .frm files carry a Begin/End control tree before the code. Those lines are
# designer output, not VBA, and their `Caption = "x"` pairs are not declarations.
code_lines <- function(path) {
  raw <- readLines(path, warn = FALSE)
  if (grepl("[.]frm$", path)) {
    start <- grep("^Option[ ]|^Attribute VB_Exposed", raw)
    if (length(start) > 0) raw <- raw[seq(min(start), length(raw))]
  }
  vapply(raw, strip_comment, character(1), USE.NAMES = FALSE)
}

source_files <- function(root) {
  f <- list.files(root, pattern = "[.](cls|bas|frm)$", recursive = TRUE,
                  full.names = TRUE)
  f[!grepl("(^|/)stale/", f)]
}

files <- source_files(if (is.null(path_arg)) "src" else path_arg)
universe_files <- source_files("src")

say("name-shadowing-scan.R -- ", length(files), " source files")

# ---------------------------------------------------------------------------
# The universe of names a variable must not take
# ---------------------------------------------------------------------------

RX_AS_TYPE <- "(?i)\\bAs\\s+(New\\s+)?([A-Za-z_][A-Za-z0-9_]*)"
RX_PUB_ENUM <- "(?i)^\\s*(Public\\s+)?Enum\\s+([A-Za-z_][A-Za-z0-9_]*)"
RX_PUB_TYPE <- "(?i)^\\s*(Public\\s+|Private\\s+)?Type\\s+([A-Za-z_][A-Za-z0-9_]*)"

captured <- function(lines, rx, group) {
  m <- regexpr(rx, lines, perl = TRUE)
  st <- attr(m, "capture.start")
  len <- attr(m, "capture.length")
  keep <- which(m > 0 & len[, group] > 0)
  substring(lines[keep], st[keep, group], st[keep, group] + len[keep, group] - 1)
}

# Every identifier written after `As`, anywhere. Built from the source so no
# hand-kept list of host types can go stale.
type_names <- character(0)
project_names <- character(0)
enum_members <- character(0)

for (f in universe_files) {
  ln <- strip_strings(code_lines(f))
  # `Open Path For Input Access Read Lock Read As File` is a file statement, not
  # a declaration, and its `As` is followed by a channel rather than a type.
  ln[grepl("(?i)^\\s*Open\\s+", ln, perl = TRUE)] <- ""
  # `As X` can appear more than once on a line (a signature), so scan all.
  mm <- gregexpr(RX_AS_TYPE, ln, perl = TRUE)
  for (i in seq_along(ln)) {
    st <- attr(mm[[i]], "capture.start")
    len <- attr(mm[[i]], "capture.length")
    if (is.null(st) || mm[[i]][1] == -1) next
    for (r in seq_len(nrow(st))) {
      type_names <- c(type_names,
                      substring(ln[i], st[r, 2], st[r, 2] + len[r, 2] - 1))
    }
  }
  project_names <- c(project_names,
                     captured(ln, RX_PUB_ENUM, 2), captured(ln, RX_PUB_TYPE, 2))

  # Enum members sit in the project namespace, not inside the enum's name.
  inside <- FALSE
  for (line in ln) {
    if (grepl(RX_PUB_ENUM, line, perl = TRUE)) { inside <- TRUE; next }
    if (grepl("(?i)^\\s*End\\s+Enum", line, perl = TRUE)) { inside <- FALSE; next }
    if (inside) {
      nm <- captured(line, "^\\s*([A-Za-z_][A-Za-z0-9_]*)", 1)
      if (length(nm) > 0) enum_members <- c(enum_members, nm)
    }
  }
}

# Components are their file names.
components <- sub("[.](cls|bas|frm)$", "", basename(universe_files))
project_names <- unique(c(components, project_names, enum_members))

type_names <- unique(type_names)
# `As` is also written before intrinsic types; those are keywords, not names a
# variable could ever take, and VBA rejects them outright.
intrinsics <- c("String", "Long", "Integer", "Double", "Boolean", "Byte",
                "Variant", "Object", "Date", "Currency", "Single", "LongLong",
                "LongPtr", "Any", "Decimal")
type_names <- type_names[!tolower(type_names) %in% tolower(intrinsics)]

forbidden <- unique(c(type_names, project_names))
forbidden <- forbidden[!tolower(forbidden) %in% tolower(intrinsics)]
forbidden_l <- tolower(forbidden)
project_l <- tolower(project_names)

# The spelling the type is written in. This project writes types in PascalCase
# and variables in camelCase, so the casing on the page is what separates
# `ShowHideLayout.Create` (meant the class) from `showHideLayout.BeginBatch`
# (meant the variable). Both are the same identifier to VBA; only one of them
# is a symptom. Project spellings win, since a component's file name is the
# authority on its own name.
canonical <- setNames(type_names, tolower(type_names))
canonical[tolower(project_names)] <- project_names

say("  ", length(forbidden), " names are types or project members")

# One name is allowed to collide, and it is written down here rather than
# quietly dropped, because a scan that always reports one finding is a scan
# people stop reading.
#
#   OBTGrantAccess.OBTGrantAccess -- a module and its only entry point, both
#   named for what the operator has to run. It is pressed with F5 in the VBE,
#   never reached by name from code, and the sandbox grant it performs is an
#   open problem in its own right (.obt/gotchas/macos-sandbox-grant.md). Renaming
#   it while that is unsolved would put a second variable into an investigation
#   that already has too many.
ALLOWED <- c("src/tests/rubberduck/OBTGrantAccess.bas:OBTGrantAccess")

# ---------------------------------------------------------------------------
# Declarations
# ---------------------------------------------------------------------------

# Every form a declared name takes. A signature declares its parameters, and the
# procedure's own name when it is a Function or Property with a return type.
RX_DIM <- paste0("(?i)^\\s*(Dim|Private|Public|Global|Static|Const)\\s+",
                 "(WithEvents\\s+)?([A-Za-z_][A-Za-z0-9_]*)")
RX_PROC <- paste0("(?i)^\\s*(Public\\s+|Private\\s+|Friend\\s+)?(Static\\s+)?",
                  "(Sub|Function|Property\\s+(Get|Let|Set))\\s+")
RX_PARAM <- paste0("(?i)(Optional\\s+)?(ByVal\\s+|ByRef\\s+)?(ParamArray\\s+)?",
                   "([A-Za-z_][A-Za-z0-9_]*)(\\(\\))?\\s+As\\s+")
RX_MEMBER <- "^\\s*([A-Za-z_][A-Za-z0-9_]*)(\\([^)]*\\))?\\s+As\\s+"

allowed_seen <- 0L
findings <- list()
add <- function(...) findings[[length(findings) + 1]] <<- list(...)

for (f in files) {
  ln <- strip_strings(code_lines(f))
  rel <- sub("^[.]/", "", f)
  is_standard_module <- grepl("[.]bas$", f)
  n <- length(ln)
  if (n == 0) next

  # Where the module's declaration section ends: the first procedure header.
  proc_starts <- which(grepl(RX_PROC, ln, perl = TRUE))

  # The procedures THIS module declares. Only these can be reached unqualified
  # from inside it, so only these can be shadowed by a local. Property Get, Let
  # and Set repeat one name, hence the unique.
  own_procs <- character(0)
  for (i in proc_starts) {
    nm <- captured(ln[i], paste0(RX_PROC, "([A-Za-z_][A-Za-z0-9_]*)"), 5)
    if (length(nm) > 0) own_procs <- c(own_procs, nm)
  }
  own_procs <- unique(own_procs)
  proc_canonical <- setNames(own_procs, tolower(own_procs))
  own_procs_l <- tolower(own_procs)
  first_proc <- if (length(proc_starts) > 0) min(proc_starts) else n + 1
  proc_ends <- which(grepl("(?i)^\\s*End\\s+(Sub|Function|Property)\\b",
                           ln, perl = TRUE))

  # Which lines sit inside a Type block, and the type's own name.
  in_type <- rep(FALSE, n)
  depth <- FALSE
  for (i in seq_len(n)) {
    if (grepl(RX_PUB_TYPE, ln[i], perl = TRUE)) { depth <- TRUE; next }
    if (grepl("(?i)^\\s*End\\s+Type", ln[i], perl = TRUE)) { depth <- FALSE; next }
    in_type[i] <- depth
  }

  # The scope a line at i belongs to, as a line range, for the read search.
  scope_of <- function(i) {
    if (i < first_proc) return(c(1L, n))
    starts <- proc_starts[proc_starts <= i]
    if (length(starts) == 0) return(c(1L, n))
    s <- max(starts)
    e <- proc_ends[proc_ends > s]
    c(s, if (length(e) > 0) min(e) else n)
  }

  # The shadow is LIVE when the scope it covers also spells the name the way the
  # TYPE is spelled and reaches through it -- `ShowHideLayout.Create`, or
  # `New ShowHideLayout`. Case is the whole test, and deliberately so: to VBA
  # both spellings are one identifier, so a search that ignored case would call
  # every ordinary `showHideLayout.BeginBatch` a symptom and bury the one line
  # that is. The declaring line itself is skipped -- it spells the type by
  # definition.
  reads_type <- function(name, from, to, skip) {
    spelling <- canonical[[tolower(name)]]
    # When the variable is spelled exactly like the type -- `Control As
    # IRibbonControl`, where MSForms.Control is the shadowed name -- casing
    # separates nothing, and every read of the variable would be counted as a
    # read of the type. `Control.ID` on a ribbon callback's own argument is not
    # a symptom of anything. The shadow is still a shadow and still reported;
    # it just cannot be called live from the page.
    if (identical(name, spelling)) return(FALSE)
    rx <- paste0("(^|[^A-Za-z0-9_.])", spelling, "\\s*\\.",
                 "|\\bNew\\s+", spelling, "\\b")
    rows <- setdiff(seq(from, to), skip)
    if (length(rows) == 0) return(FALSE)
    any(grepl(rx, ln[rows], perl = TRUE))
  }

  # The shadow is LIVE when the scope it covers calls the procedure by the
  # spelling the PROCEDURE is written in. Casing is the test, for the same
  # reason it is the test for types: to VBA the two spellings are one
  # identifier, so ignoring case would call every ordinary read of the variable
  # a symptom. A call is any unqualified mention -- `Foo x, y` as a statement,
  # `Foo(x)` as a function, `Call Foo`, `Then Foo` -- so the pattern asks only
  # that the name stand alone and not follow a dot.
  calls_proc <- function(name, from, to, skip) {
    spelling <- proc_canonical[[tolower(name)]]
    if (is.null(spelling)) return(FALSE)
    # Spelled exactly like the procedure: casing separates nothing and every
    # read of the variable would count. Still a shadow, just not callable live
    # from the page.
    if (identical(name, spelling)) return(FALSE)
    header_rx <- paste0(RX_PROC, spelling, "\\b")
    call_rx <- paste0("(^|[^A-Za-z0-9_.])", spelling, "\\s*(\\(|,|\\s|$)")
    rows <- setdiff(seq(from, to), skip)
    # The procedure's own header spells its name by definition.
    if (length(rows) > 0) rows <- rows[!grepl(header_rx, ln[rows], perl = TRUE)]
    if (length(rows) == 0) return(FALSE)
    any(grepl(call_rx, ln[rows], perl = TRUE))
  }

  report <- function(i, name, kind) {
    key <- tolower(name)
    in_types <- key %in% forbidden_l
    # A procedure cannot shadow itself, and a Type field is scoped to its type.
    in_procs <- (key %in% own_procs_l) && !(kind %in% c("procedure", "member"))
    if (!in_types && !in_procs) return()
    if (paste0(rel, ":", name) %in% ALLOWED) {
      allowed_seen <<- allowed_seen + 1L
      return()
    }

    sc <- if (kind == "member") NULL else scope_of(i)

    # A live procedure shadow is reported ahead of any type verdict: the call
    # that is quietly not happening is the worse of the two faults, and it is
    # the one that names the fix.
    if (in_procs && !is.null(sc) && calls_proc(name, sc[1], sc[2], i)) {
      add(file = rel, line = i, name = name, kind = kind, sev = "calling",
          text = trimws(ln[i]))
      return()
    }

    if (in_types) {
      sev <- if (kind == "member") "member"
             else if (key %in% project_l) "error" else "builtin"
      if (sev != "member") {
        if (reads_type(name, sc[1], sc[2], i)) sev <- "breaking"
      }
      add(file = rel, line = i, name = name, kind = kind, sev = sev,
          text = trimws(ln[i]))
      return()
    }

    add(file = rel, line = i, name = name, kind = kind, sev = "procname",
        text = trimws(ln[i]))
  }

  for (i in seq_len(n)) {
    line <- ln[i]
    if (trimws(line) == "") next

    if (in_type[i]) {
      nm <- captured(line, RX_MEMBER, 1)
      if (length(nm) > 0) report(i, nm, "member")
      next
    }

    # Procedures BEFORE declarations. `Private Sub Foo(ByVal x As Y)` opens with
    # `Private ` and would otherwise be read as a declaration of a variable
    # called `Sub`, and its parameter list never looked at.
    if (grepl(RX_PROC, line, perl = TRUE)) {
      # A procedure's own name is a declaration too, but only in a standard
      # module is it a PROJECT-level identifier. A class member is reached
      # through an object, so `Public Property Get Name()` on a class is the
      # ordinary shape of a property and shadows nothing outside its own body --
      # thirteen classes here spell it that way, plus Axis, Series and Workbook.
      # Reporting those would bury the standard-module cases that do matter.
      if (is_standard_module) {
        pname <- captured(line, paste0(RX_PROC, "([A-Za-z_][A-Za-z0-9_]*)"), 5)
        if (length(pname) > 0) report(i, pname, "procedure")
      }

      # Parameters, wherever the signature continues onto further lines.
      j <- i
      sig <- line
      while (grepl("_\\s*$", sig, perl = TRUE) && j < n) {
        j <- j + 1
        sig <- paste(sub("_\\s*$", "", sig), ln[j])
      }
      inner <- sub("^[^(]*\\(", "", sig)
      inner <- sub("\\)[^)]*$", "", inner)
      if (nzchar(trimws(inner))) {
        for (part in strsplit(inner, ",", fixed = TRUE)[[1]]) {
          nm <- captured(part, RX_PARAM, 4)
          if (length(nm) > 0) report(i, nm, "parameter")
        }
      }
      next
    }

    if (grepl(RX_DIM, line, perl = TRUE)) {
      # `Dim a As X, b As Y` declares both.
      body <- sub(RX_DIM, "\\3", line, perl = TRUE)
      first <- captured(line, RX_DIM, 3)
      if (length(first) > 0) report(i, first, "variable")
      rest <- regmatches(line, gregexpr(
        ",\\s*([A-Za-z_][A-Za-z0-9_]*)\\s+As\\s+", line, perl = TRUE))[[1]]
      for (r in rest) {
        nm <- sub("^,\\s*", "", sub("(?i)\\s+As\\s+$", "", r, perl = TRUE))
        report(i, nm, "variable")
      }
      next
    }
  }
}

# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

order_of <- c(breaking = 1, calling = 2, error = 3, procname = 4,
              builtin = 5, member = 6)
if (length(findings) > 0) {
  sev <- vapply(findings, function(x) x$sev, character(1))
  findings <- findings[order(order_of[sev],
                             vapply(findings, function(x) x$file, character(1)),
                             vapply(findings, function(x) x$line, numeric(1)))]
}

# --breaking means "the live ones". A shadowed procedure that is actually
# called is as live as a shadowed type that is actually read.
if (only_breaking) {
  findings <- Filter(function(x) x$sev %in% c("breaking", "calling"), findings)
}

counts <- table(vapply(findings, function(x) x$sev, character(1)))

if (length(findings) == 0) {
  say("")
  say("No variable is named after a type or after a procedure of its module.",
      "  (", allowed_seen, " allowed collision(s) skipped)")
  quit(status = 0)
}

headline <- c(
  breaking = "BREAKING -- the shadowed TYPE is read in this scope, so the read reaches the variable (error 91 at run time)",
  calling  = "CALLING -- the shadowed PROCEDURE is called in this scope, so the call reaches the variable and the procedure never runs",
  error    = "ERROR -- shadows a name this project defines",
  procname = "PROCNAME -- shadows a procedure of the same module, with no call to it found in this scope",
  builtin  = "BUILTIN -- shadows a host type",
  member   = "MEMBER -- a Type field named after its own type; scoped, so it works, and it is still banned"
)

cat("\n")
last <- ""
for (x in findings) {
  if (x$sev != last) {
    cat("\n", headline[[x$sev]], "\n",
        strrep("-", 79), "\n", sep = "")
    last <- x$sev
  }
  cat(sprintf("  %s:%d\n    %s  (%s `%s`)\n",
              x$file, x$line, x$text, x$kind, x$name))
}

cat("\n", strrep("=", 79), "\n", sep = "")
cat("findings: ",
    paste(sprintf("%s=%d", names(counts), as.integer(counts)), collapse = "  "),
    "\n", sep = "")

# What actually gates. Without --strict, a run whose only findings are procname
# is reported and passes: nothing there is broken, and failing on all of them
# would retire the scan as a pre-flight the day this check was added.
gating <- Filter(function(x) strict || x$sev != "procname", findings)
if (length(gating) == 0) {
  cat("gate: pass -- every finding is procname, which is latent. ",
      "Run with --strict to gate on those too.\n", sep = "")
  quit(status = 0)
}
quit(status = 1)

#!/usr/bin/env Rscript
#
# linelist-inspect.R -- judge a delivered linelist without opening Excel.
#
# Usage:
#   Rscript scripts/devtools/linelist-inspect.R <linelist.xlsb>
#   Rscript scripts/devtools/linelist-inspect.R <linelist.xlsb> --quiet
#
# Exit code 1 when a check fails, 0 when the delivered file is clean.
#
# WHY THIS EXISTS
# -----------------------------------------------------------------------------
# A generation reports "OK" over a linelist that does not compile, is written
# entirely in tags and carries no geography. Every one of those was found by
# opening the delivered file in Excel and noticing. This reads the bytes
# instead: an .xlsb is a zip, and the three parts worth reading are
# docProps/app.xml (the worksheet names), xl/workbook.bin (the defined names)
# and xl/vbaProject.bin (the VBA components).
#
# It is the AFTER half of a before/after. Run it on the file a build produced
# AND on the one before it: a check that reports the fault on the old file and
# passes on the new is evidence. A check that only ever passes is not.
#
# WHAT IT READS, AND WHAT IT DELIBERATELY DOES NOT
# -----------------------------------------------------------------------------
# Component names come from the PROJECT stream inside vbaProject.bin, which is
# plain text -- `Module=`, `Class=`, `BaseClass=`, `Document=` lines -- and
# needs no decoding at all.
#
# The module SOURCE does: it is compressed with the MS-OVBA run-length scheme,
# and decoding that here would be a decompressor written in R to answer
# questions the repository can already answer from src/. So the ribbon check
# below compares the delivered customUI against the SOURCE of
# EventsLinelistRibbon rather than against the delivered copy of it. That is
# the drift worth catching -- a button bound to a callback nobody wrote -- and
# it costs nothing.
#
# WHAT IT CHECKS
# -----------------------------------------------------------------------------
# 1. The three classes that must travel, and the modules that must not.
#    SectionMap, GeoFormCache and AnalysisRanges are named as TYPES and as
#    predeclared factories, never by a bare call, which is why the call scan
#    never saw them missing. Any FormLogic* present is a fault by itself: those
#    modules use Me and declare control event handlers, and neither compiles in
#    a standard module.
# 2. The empty starting worksheet, judged by NAME PATTERN across the languages
#    this project ships in, because Excel calls it Sheet1 or Feuille1 depending
#    on the host.
# 3. Translation, through two independent symptoms of one fault: worksheets
#    still called LLSHEET_*, and RNG_LLLanguageCode holding something that is
#    not a column of the translation tables. A language NAME where a CODE
#    belongs resolves no column, and every lookup then falls back to the tag.
# 4. Ribbon callbacks. Every onAction in the delivered customUI has to name a
#    Public Sub of EventsLinelistRibbon, and that module must not carry
#    Option Private Module, which takes its members out of name dispatch.
# 5. Geo expansion. One `geo` control variable becomes twelve dictionary rows:
#    four admin levels, four p-codes, four concatenations. pcode_adm* and
#    concat_adm* names exist only if AppendGeoLines ran, so they are the
#    signal. adm1_concat..adm4_concat are the geo SHEET's own lookup tables and
#    prove nothing either way, so they are excluded.
#
# WHAT IT CANNOT SEE
# -----------------------------------------------------------------------------
# Whether the project COMPILES. That needs a VBA host. This reports the two
# compile faults it can recognise by name and nothing else, so a clean report
# here is necessary rather than sufficient.

args <- commandArgs(trailingOnly = TRUE)
quiet <- "--quiet" %in% args
target <- args[!grepl("^--", args)][1]

repo_root <- normalizePath(file.path(dirname(sub("^--file=", "", grep(
  "^--file=", commandArgs(trailingOnly = FALSE), value = TRUE
)[1])), "..", ".."))

if (is.na(target) || !file.exists(target)) {
  cat("Usage: Rscript scripts/devtools/linelist-inspect.R <linelist.xlsb>\n")
  quit(status = 2)
}

REQUIRED_CLASSES <- c("SectionMap", "GeoFormCache", "AnalysisRanges")
LANGUAGE_CODES <- c("ENG", "FRA", "SPA", "POR", "ARA")

# Excel's placeholder worksheet, in the languages this project ships in. The
# name is localised, which is exactly why it survived a template build unnoticed.
PLACEHOLDER <- "^(Sheet|Feuille|Hoja|Folha|Blatt|Foglio)[ ]*[0-9]+$"

# ---------------------------------------------------------------------------
# Reading parts out of the workbook
# ---------------------------------------------------------------------------

work_dir <- file.path(tempdir(), "obt-linelist-inspect")
unlink(work_dir, recursive = TRUE)
dir.create(work_dir, recursive = TRUE, showWarnings = FALSE)

part_names <- unzip(target, list = TRUE)$Name

read_part <- function(part) {
  if (!(part %in% part_names)) {
    return(NULL)
  }
  unzip(target, files = part, exdir = work_dir, overwrite = TRUE)
  path <- file.path(work_dir, part)
  readBin(path, "raw", n = file.info(path)$size)
}

# Printable ASCII runs of a binary part. A non-printable byte becomes a newline
# rather than vanishing: dropping it would JOIN the two runs either side and
# "Module=Foo" followed by "Class=Bar" would read as one token.
ascii_tokens <- function(bytes) {
  if (is.null(bytes)) {
    return(character(0))
  }
  keep <- bytes >= as.raw(32) & bytes <= as.raw(126)
  chars <- rep("\n", length(bytes))
  chars[keep] <- rawToChar(bytes[keep], multiple = TRUE)
  unlist(strsplit(paste(chars, collapse = ""), "\n+"))
}

# The same idea over UTF-16LE, which is how workbook.bin stores defined names.
# Both byte alignments are scanned because a run may start at either, and the
# two are kept APART rather than concatenated: token order only means anything
# within one alignment, and reading across the seam once made "the value after
# RNG_LLLanguageCode" a name from the other pass entirely.
utf16_alignments <- function(bytes) {
  if (is.null(bytes)) {
    return(list(character(0)))
  }
  lapply(1:2, function(start) {
    idx <- seq(start, length(bytes) - 1L, by = 2L)
    lo <- bytes[idx]
    hi <- bytes[idx + 1L]
    ok <- hi == as.raw(0) & lo >= as.raw(32) & lo <= as.raw(126)
    chars <- rep("\n", length(idx))
    chars[ok] <- rawToChar(lo[ok], multiple = TRUE)
    unlist(strsplit(paste(chars, collapse = ""), "\n+"))
  })
}

# The value stored right after a hidden name, found by BYTE OFFSET.
#
# Adjacency in a token list is not reliable here: the token stream depends on
# which byte a UTF-16 run is assumed to start at, and reading "the token after
# RNG_LLLanguageCode" out of the wrong alignment answered a neighbouring name
# instead of the value. So the key is located as raw UTF-16LE bytes and only
# the window immediately behind it is decoded, at the alignment the key itself
# establishes. That leaves no guesswork.
value_after <- function(bytes, key, window = 160L) {
  if (is.null(bytes)) {
    return(NA_character_)
  }
  key_bytes <- as.raw(as.vector(rbind(utf8ToInt(key), 0L)))
  at <- grepRaw(key_bytes, bytes, all = TRUE)
  if (!length(at)) {
    return(NA_character_)
  }

  start <- at[1] + length(key_bytes)
  stop <- min(start + window, length(bytes) - 1L)
  idx <- seq(start, stop, by = 2L)
  lo <- bytes[idx]
  hi <- bytes[idx + 1L]
  ok <- hi == as.raw(0) & lo >= as.raw(32) & lo <= as.raw(126)
  chars <- rep("\n", length(idx))
  chars[ok] <- rawToChar(lo[ok], multiple = TRUE)
  tokens <- unlist(strsplit(paste(chars, collapse = ""), "\n+"))

  # The type marker that follows the name is not the value, and neither is the
  # next hidden name along.
  tokens <- tokens[!grepl("^(RNG_|HiddenNames:)", tokens)]
  tokens <- tokens[grepl("^[A-Za-z][A-Za-z0-9 _-]*$", tokens)]
  if (!length(tokens)) NA_character_ else tokens[1]
}

worksheet_names <- function() {
  bytes <- read_part("docProps/app.xml")
  if (is.null(bytes)) {
    return(character(0))
  }
  xml <- rawToChar(bytes)
  Encoding(xml) <- "UTF-8"

  count <- regmatches(xml, regexec(
    "<vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>([0-9]+)", xml
  ))[[1]]
  if (length(count) < 2) {
    return(character(0))
  }

  # HeadingPairs uses <vt:lpstr> for its own labels, so the names are taken out
  # of TitlesOfParts alone -- otherwise the first two are "Worksheets" and
  # "Named Ranges".
  titles <- regmatches(xml, regexec("<TitlesOfParts>(.*?)</TitlesOfParts>", xml))[[1]]
  if (length(titles) < 2) {
    return(character(0))
  }
  all_names <- regmatches(titles[2], gregexpr("<vt:lpstr>(.*?)</vt:lpstr>", titles[2]))[[1]]
  all_names <- gsub("</?vt:lpstr>", "", all_names)
  head(all_names, as.integer(count[2]))
}

vba_components <- function() {
  tokens <- ascii_tokens(read_part("xl/vbaProject.bin"))
  hits <- grep("^(Module|Class|BaseClass|Document)=[A-Za-z0-9_]+$", tokens, value = TRUE)
  kinds <- sub("=.*$", "", hits)
  names_of <- sub("^[^=]+=", "", hits)
  split(unique(names_of), kinds[!duplicated(names_of)])
}

customui_actions <- function() {
  part <- grep("customUI(14)?\\.xml$", part_names, value = TRUE, ignore.case = TRUE)
  if (!length(part)) {
    return(NULL)
  }
  xml <- rawToChar(read_part(part[1]))
  hits <- regmatches(xml, gregexpr("onAction[ ]*=[ ]*\"[^\"]+\"", xml))[[1]]
  sort(unique(gsub("^onAction[ ]*=[ ]*\"|\"$", "", hits)))
}

# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------

failed <- FALSE

check <- function(ok, label, detail = "") {
  if (!ok) failed <<- TRUE
  cat(sprintf("   [%4s] %-44s %s\n", if (ok) "ok" else "FAIL", label, detail))
}

note <- function(label, detail = "") {
  cat(sprintf("   [  --] %-44s %s\n", label, detail))
}

sheets <- worksheet_names()
book_bytes <- read_part("xl/workbook.bin")
names_in_book <- unlist(utf16_alignments(book_bytes))
components <- vba_components()
code <- sort(unique(c(components$Module, components$Class, components$BaseClass)))

cat(strrep("=", 74), "\n", sep = "")
cat(basename(target), "\n")
cat(strrep("=", 74), "\n", sep = "")

if (!length(code)) {
  cat("\nThis workbook carries no VBA project at all.\n")
  quit(status = 1)
}

cat("\n1. components the project must and must not carry\n")
for (name in REQUIRED_CLASSES) {
  present <- name %in% code
  check(present, name,
        if (present) "transferred" else "MISSING -> the linelist cannot compile")
}
stray <- grep("^FormLogic", code, value = TRUE)
check(!length(stray), "no FormLogic module",
      if (!length(stray)) "none present" else
        paste0("PRESENT -> ", paste(stray, collapse = ", "),
               " (they use Me; a standard module cannot)"))

cat("\n2. the empty starting worksheet\n")
placeholder <- grep(PLACEHOLDER, sheets, value = TRUE)
check(!length(placeholder), "placeholder sheet removed",
      if (!length(placeholder)) "gone" else
        paste0("still there -> ", paste(placeholder, collapse = ", ")))

cat("\n3. translation\n")
tagged <- grep("^LLSHEET_", sheets, value = TRUE)
check(!length(tagged), "worksheets carry translated names",
      if (!length(tagged)) "translated" else
        sprintf("%d still named by tag -> %s", length(tagged),
                paste(head(tagged, 3), collapse = ", ")))

# The value sits a token or two after the name, not always immediately after:
# the record carries a type marker between them, and whether that marker
# survives the ASCII filter depends on its bytes. So the next few tokens are
# scanned and the first that looks like a value wins. Taking the very next
# token reported '@' for a workbook whose code was ENG -- a false failure,
# which is the one kind of answer a check like this must never give.
lang_code <- value_after(book_bytes, "RNG_LLLanguageCode")
# A hidden name and its value are not always stored next to each other -- how
# far apart depends on the record -- so this cannot always be read out. When it
# cannot, that is reported as UNDETERMINED rather than as a failure: a check
# that answers "broken" when it means "I could not tell" is worse than no check,
# and the worksheet names above already carry the same fault where it shows.
if (is.na(lang_code)) {
  note("RNG_LLLanguageCode not readable here",
       "value not stored beside the name; see the worksheet names above")
} else {
  check(lang_code %in% LANGUAGE_CODES, "RNG_LLLanguageCode is a real column",
        if (lang_code %in% LANGUAGE_CODES) sprintf("'%s'", lang_code) else
          sprintf("'%s' is not one of %s", lang_code,
                  paste(LANGUAGE_CODES, collapse = ", ")))
}

cat("\n4. ribbon callbacks resolve\n")
actions <- customui_actions()
ribbon_source <- file.path(repo_root, "src/modules/linelist/EventsLinelistRibbon.bas")

if (is.null(actions)) {
  note("no customUI in this workbook", "the buttons build carries none")
} else if (!file.exists(ribbon_source)) {
  note("EventsLinelistRibbon.bas not found", "cannot resolve the bindings")
} else {
  source_text <- readLines(ribbon_source, warn = FALSE)
  defined <- regmatches(source_text, regexec(
    "^[ \t]*(?:Public[ ]+)?Sub[ ]+([A-Za-z_][A-Za-z0-9_]*)", source_text
  ))
  defined <- unlist(lapply(defined, function(m) if (length(m) >= 2) m[2] else NULL))

  missing <- setdiff(actions, defined)
  check(!length(missing), sprintf("%d onAction bindings defined", length(actions)),
        if (!length(missing)) "all resolve" else
          paste0("undefined -> ", paste(missing, collapse = ", ")))

  check("EventsLinelistRibbon" %in% code, "the ribbon module was transferred",
        if ("EventsLinelistRibbon" %in% code) "present" else "MISSING")

  private <- any(grepl("^[ \t]*Option[ ]+Private[ ]+Module", source_text))
  check(!private, "callback module is not Option Private Module",
        if (!private) "plain module" else
          "PRIVATE -> its members are hidden from name dispatch")
}

cat("\n5. geo variables expanded into their twelve rows\n")
# The geo sheet's own lookup tables, there whether or not a variable ever
# expanded, so they are not evidence either way.
sheet_tables <- c(paste0("adm", 1:4, "_concat"), "hf_concat")
admin <- setdiff(unique(grep("^adm[1-4]_", names_in_book, value = TRUE)), sheet_tables)
pcode <- unique(grep("^pcode_adm[1-4]_", names_in_book, value = TRUE))
concat <- unique(grep("^concat_adm[1-4]_", names_in_book, value = TRUE))

cat(sprintf("          adm*=%d  pcode_adm*=%d  concat_adm*=%d\n",
            length(admin), length(pcode), length(concat)))

if (!length(admin) && !length(pcode) && !length(concat)) {
  note("no geo variable in this build", "nothing to expand")
} else {
  check(length(pcode) >= 4 && length(concat) >= 4, "AppendGeoLines ran",
        if (length(pcode) >= 4) "expanded" else
          "geo variables stayed one column each")
}

if (!quiet) {
  cat(sprintf("\n   %d worksheets, %d code components\n", length(sheets), length(code)))
  cat("   ", paste(sheets, collapse = ", "), "\n", sep = "")
}

unlink(work_dir, recursive = TRUE)
quit(status = if (failed) 1 else 0)

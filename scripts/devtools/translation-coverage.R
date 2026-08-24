#!/usr/bin/env Rscript
#
# translation-coverage.R -- what the translation workbook still owes, and what
# it carries that nothing asks for any more.
#
# trads/designer_translations.xlsx holds eight tables. Four of them dress the
# linelist the designer generates, four dress the designer workbook itself:
#
#   LinelistTranslation     t_tradllshapes   button captions on the sheets
#                           t_tradllmsg      messages, sheet names, dropdowns
#                           t_tradllforms    form captions and control captions
#                           t_tradllribbon   the linelist ribbon labels
#
#   DesignerTranslation     t_tradmsg        messages AND designer ribbon labels
#                           t_tradrange      the labels on the main worksheet
#                           t_tradshape      the buttons on the main worksheet
#                           t_traddrop       the dropdowns of the main worksheet
#
# Every row is keyed by a tag. A tag the code asks for and the table does not
# carry comes back as the tag itself: TranslatedValue returns the text it was
# handed when the lookup misses, so the user reads "MSG_ErrUpdate" on screen and
# nothing raises. A row nobody asks for costs nothing at run time but is five
# languages of upkeep, and it hides in a table of 175 rows.
#
# This script reports both, per table.
#
# WHERE THE ASKED-FOR TAGS COME FROM
# -----------------------------------------------------------------------------
# Four readers, because the eight tables are keyed on four different things.
#
#   1. VBA source     src/classes and src/modules, comments stripped. A tag is
#                     a string literal shaped like one (MSG_, MSGB_, LLSHEET_,
#                     INSTSHEET, SHP_, RNG_, DROP...) or any literal handed to
#                     TranslatedValue or ValueExists. src/tests and
#                     src/classes/stale are left out: a tag only a test seeds
#                     is not a tag the product asks for.
#
#   2. Form binaries  the control names inside the .frx files. TranslateForm
#                     walks UserForm.Controls and looks each control up BY ITS
#                     NAME, so the names in the binary are the tags. Only the
#                     six control types it translates are read -- CommandButton,
#                     Label, OptionButton, Page, Frame, CheckBox -- which in
#                     this project are the CMD_, LBL_, OPT_, PGE_, FRM_ and
#                     CHK_ prefixes. A ListBox or a TextBox has no Caption and
#                     is skipped by the same routine, so LST_ and TXT_ are not
#                     tags. The form's own name is one too, but only for the
#                     nine forms whose FormLogic module runs
#                     `Me.Caption = TranslatedValue(Me.Name)`. The other two
#                     keep their design-time caption in every language.
#
#                     The .frx is a binary and the names sit in it as
#                     length-prefixed strings. A run of printable bytes is not
#                     enough on its own -- names are padded to a four-byte
#                     boundary with junk, so a plain string scan reads
#                     CMD_NewKey as "CMD_NewKeyi". The length marker is read
#                     instead, which gets the name exactly.
#
#   3. Ribbon XML     a control carries getLabel="LangLabel" when its label is
#                     fetched from a translation table at load, and a plain
#                     label="..." when it is not. Only the getLabel ones are
#                     tags. The linelist reads the two _ribbontemplate folders,
#                     the designer reads its own.
#
#   4. The designer   src/bin/designer/designer.xlsb, for the two tables keyed
#      binary         on things that live in the file rather than in the code:
#                     t_tradrange on defined names, t_tradshape on shape names.
#                     The binary is gitignored and travels through the release
#                     asset store, so it may be absent. When it is, those two
#                     tables are reported as unchecked rather than as dead.
#
# WHAT IS A HINT AND WHAT IS NOT
# -----------------------------------------------------------------------------
# The message tables are the one place the answer is softer. A message tag is a
# string literal and nothing in the literal says which of the two workbooks
# should carry it, so a tag missing from BOTH message tables is reported once,
# with the files that use it and a designer/linelist guess read off the folder
# those files sit in. The guess is a guess; the file list under it is not.
#
# Everything else is exact: a form control name belongs to a linelist form, a
# ribbon id belongs to the ribbon it was read from, a shape name belongs to the
# designer worksheet.
#
# ONE THING THIS CANNOT SEE
# -----------------------------------------------------------------------------
# The designer workbook's own document modules -- ThisWorkbook and the sheet
# modules -- live inside designer.xlsb and nowhere in src. Every ribbon callback
# and every designer class IS in src, so a tag reached from one of those is
# read; a tag a document module asks for on its own is not, and would read as a
# dead row here. Check a designer row against the file before deleting it.
#
# Usage:
#   Rscript scripts/devtools/translation-coverage.R
#   Rscript scripts/devtools/translation-coverage.R --out trads/report.md
#   Rscript scripts/devtools/translation-coverage.R --workbook <path.xlsx>
#   Rscript scripts/devtools/translation-coverage.R --quiet
#
# Writes trads/translation-coverage.md by default. Exit code 1 when anything is
# missing or dead, 0 when the eight tables match what the product asks for.

suppressMessages({
  library(openxlsx)
  library(readxl)
})

# ---------------------------------------------------------------------------
# Arguments and repository root
# ---------------------------------------------------------------------------

args <- commandArgs(trailingOnly = TRUE)
quiet <- "--quiet" %in% args

arg_value <- function(flag, fallback) {
  hit <- match(flag, args)
  if (is.na(hit) || hit == length(args)) {
    return(fallback)
  }
  args[hit + 1L]
}

script_path <- sub(
  "^--file=",
  "",
  grep("^--file=", commandArgs(trailingOnly = FALSE), value = TRUE)[1]
)
repo_root <- normalizePath(file.path(dirname(script_path), "..", ".."))
setwd(repo_root)

workbook_path <- arg_value("--workbook", "trads/designer_translations.xlsx")
out_path <- arg_value("--out", "trads/translation-coverage.md")

if (!file.exists(workbook_path)) {
  stop(
    "translation workbook not found: ",
    workbook_path,
    "\nIt is gitignored. Pull it with scripts/release/pull-assets.sh, or ",
    "point --workbook at a copy.",
    call. = FALSE
  )
}

say <- function(...) {
  if (!quiet) {
    cat(..., "\n", sep = "")
  }
}

# ---------------------------------------------------------------------------
# Reading VBA source
# ---------------------------------------------------------------------------

# A quote inside a VBA string is doubled, so one pass that flips a flag on every
# quote gets the boundaries right. Everything after an unquoted apostrophe is a
# comment, and a tag named in a comment is not a tag anybody asks for.
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

# The shapes a translation tag takes in this project. RNG_ and the bare
# dropdown names are here because the designer tables are keyed on them too.
TAG_PATTERN <- paste0(
  "^(",
  "MSG_[A-Za-z0-9_]+",
  "|MSGB_[A-Za-z0-9_]+",
  "|LLSHEET_[A-Za-z0-9_]+",
  "|INSTSHEET[A-Za-z0-9_]*",
  "|SHP_[A-Za-z0-9_]+",
  "|NoteText_[A-Za-z0-9_]+",
  ")$"
)

# Every literal the file hands to a lookup, plus every literal shaped like a
# tag. The first catches a tag whose name breaks the convention, the second
# catches the Private Const declarations that never sit at a call site.
tags_in_file <- function(path) {
  raw <- readLines(path, warn = FALSE, encoding = "UTF-8")
  code <- vapply(raw, strip_comment, character(1), USE.NAMES = FALSE)

  literals <- regmatches(code, gregexpr("\"[^\"]*\"", code))
  looked_up <- regmatches(
    code,
    gregexpr(
      "(Translate[A-Za-z]*|ValueExists|IsTranslated)\\([ ]*\"[^\"]*\"",
      code
    )
  )

  flat_literals <- gsub("^\"|\"$", "", unlist(literals))
  flat_looked_up <- sub(
    ".*\"([^\"]*)\"$",
    "\\1",
    unlist(looked_up)
  )

  shaped <- flat_literals[grepl(TAG_PATTERN, flat_literals)]
  unique(c(shaped, flat_looked_up[nzchar(flat_looked_up)]))
}

# The workbook a folder's code ends up in. Only a hint, and only used to guess
# which sheet a missing message tag belongs on.
side_of <- function(path) {
  if (grepl("/(classes|modules)/(designer|dev)/|/modules/headless/", path)) {
    return("designer")
  }
  if (
    grepl(
      paste0(
        "/classes/(linelist|dataio|analyses|sections|showhide|geo|graphs|",
        "formulas)/|/modules/(linelist|linelistform)/"
      ),
      path
    )
  ) {
    return("linelist")
  }
  if (grepl("/(classes|modules)/(setup|mastersetup|msetup)/", path)) {
    return("setup")
  }
  "shared"
}

read_source_tags <- function() {
  files <- list.files(
    c("src/classes", "src/modules"),
    pattern = "\\.(bas|cls)$",
    recursive = TRUE,
    full.names = TRUE
  )
  files <- files[!grepl("/stale/", files, fixed = TRUE)]

  rows <- lapply(files, function(path) {
    found <- tags_in_file(path)
    if (!length(found)) {
      return(NULL)
    }
    data.frame(
      tag = found,
      file = path,
      side = side_of(path),
      stringsAsFactors = FALSE
    )
  })
  rows <- Filter(Negate(is.null), rows)
  if (!length(rows)) {
    return(data.frame(
      tag = character(),
      file = character(),
      side = character(),
      stringsAsFactors = FALSE
    ))
  }
  do.call(rbind, rows)
}

# ---------------------------------------------------------------------------
# Reading the form binaries
# ---------------------------------------------------------------------------

# The .frx stores a control name as a two-byte length with the high bit of the
# fourth byte set, then a fixed run of record fields, then the name. The run is
# 16 bytes long for some control records and 20 for others, so each candidate
# offset is tried and the first one that spells an identifier of exactly the
# stated length wins. Anything else in the file fails that test.
FRX_NAME_OFFSETS <- c(16L, 20L, 24L, 12L, 28L, 32L)

# A-Z, a-z, 0-9 and the underscore, as byte values.
is_identifier_byte <- function(values) {
  (values >= 48L & values <= 57L) |
    (values >= 65L & values <= 90L) |
    (values >= 97L & values <= 122L) |
    values == 95L
}

frx_names <- function(path) {
  bytes <- readBin(path, "raw", file.info(path)$size)
  values <- as.integer(bytes)
  total <- length(values)
  if (total < 64L) {
    return(character())
  }

  limit <- total - max(FRX_NAME_OFFSETS) - 64L
  if (limit < 1L) {
    return(character())
  }
  starts <- seq_len(limit)
  markers <- starts[values[starts + 2L] == 0L & values[starts + 3L] == 128L]

  text_found <- character()
  first_byte <- integer()
  last_byte <- integer()

  for (start in markers) {
    len <- values[start] + 256L * values[start + 1L]
    if (len < 1L || len > 64L) {
      next
    }
    for (offset in FRX_NAME_OFFSETS) {
      first <- start + offset
      last <- first + len - 1L
      if (last > total) {
        next
      }
      slice <- values[first:last]
      # rawToChar refuses a slice holding a zero byte, and a name never does.
      if (!all(is_identifier_byte(slice))) {
        next
      }
      text <- rawToChar(as.raw(slice))
      if (grepl("^[A-Za-z_][A-Za-z0-9_]*$", text)) {
        text_found <- c(text_found, text)
        first_byte <- c(first_byte, first)
        last_byte <- c(last_byte, last)
        break
      }
    }
  }

  if (!length(text_found)) {
    return(character())
  }

  # A marker whose real field is something else can still land inside a name
  # that another marker reads in full, which is how OptionButton1 also comes
  # back as OptionB. A hit whose bytes sit wholly inside a longer hit is that,
  # and never a control of its own.
  inside <- vapply(
    seq_along(text_found),
    function(idx) {
      any(
        first_byte <= first_byte[idx] &
          last_byte >= last_byte[idx] &
          (last_byte - first_byte) > (last_byte[idx] - first_byte[idx])
      )
    },
    logical(1)
  )

  unique(text_found[!inside])
}

# The six control types TranslateForm knows, as this project names them.
FORM_TAG_PATTERN <- "^(CMD|CHK|LBL|OPT|PGE|FRM)_[A-Za-z0-9_]+$|^Option[A-Z]"

# A control nobody renamed. Its caption is never translated, because no table
# will ever carry a row called Label1.
DEFAULT_NAME_PATTERN <- paste0(
  "^(Label|CommandButton|OptionButton|CheckBox|Frame|Page|ToggleButton)",
  "[0-9]+$"
)

# The forms are gitignored: they are exported next to a build or a test run.
# Whichever folder is on disk is read; when neither is, the form table goes
# unchecked.
FORM_FOLDERS <- c(".mock/forms/designer", ".test-runner/forms/merged")

# The form's own name is a tag only where the code asks for it. A FormLogic
# module that runs `Me.Caption = TranslatedValue(Me.Name)` reads the row keyed
# on its form name; one that does not leaves the form with its design-time
# caption in every language, and that row is a row nothing reads.
#
# The form name is read from the module's own @ModuleDescription, because two
# file names do not carry it: FormLogicShowHide belongs to F_ShowHideLL, and
# FormLogicEpiWeek's description names no form, so that one falls back to the
# file name. Deriving every name from the file name reported F_ShowHideLL as a
# dead row while its caption was being translated.
# scripts/headless/merge-form-code.R states the same pairing.
form_of_module <- function(path) {
  lines <- readLines(path, warn = FALSE)
  described <- grep("@ModuleDescription", lines, value = TRUE)
  named <- regmatches(described, regexpr("F_[A-Za-z0-9_]+", described))
  if (length(named)) {
    return(named[1])
  }
  paste0("F_", sub("^FormLogic", "", sub("\\.bas$", "", basename(path))))
}

forms_that_translate_their_caption <- function() {
  modules <- list.files(
    "src/modules/linelistform",
    pattern = "^FormLogic.*\\.bas$",
    full.names = TRUE
  )
  asking <- Filter(
    function(path) {
      text <- paste(readLines(path, warn = FALSE), collapse = "\n")
      grepl("TranslatedValue\\([ ]*Me\\.Name[ ]*\\)", text)
    },
    modules
  )
  sort(unique(vapply(asking, form_of_module, character(1), USE.NAMES = FALSE)))
}

read_form_tags <- function() {
  folder <- FORM_FOLDERS[dir.exists(FORM_FOLDERS)][1]
  if (is.na(folder)) {
    return(NULL)
  }

  binaries <- list.files(folder, pattern = "\\.frx$", full.names = TRUE)
  if (!length(binaries)) {
    return(NULL)
  }

  captioned <- forms_that_translate_their_caption()

  rows <- lapply(binaries, function(path) {
    form <- sub("\\.frx$", "", basename(path))
    names_found <- frx_names(path)
    defaults <- names_found[grepl(DEFAULT_NAME_PATTERN, names_found)]
    controls <- setdiff(
      names_found[grepl(FORM_TAG_PATTERN, names_found)],
      defaults
    )
    if (form %in% captioned) {
      controls <- c(form, controls)
    }
    if (!length(controls) && !length(defaults)) {
      return(NULL)
    }
    data.frame(
      tag = c(controls, defaults),
      form = form,
      unnamed = c(rep(FALSE, length(controls)), rep(TRUE, length(defaults))),
      stringsAsFactors = FALSE
    )
  })

  rows <- Filter(Negate(is.null), rows)
  if (!length(rows)) {
    return(NULL)
  }

  list(folder = folder, rows = do.call(rbind, rows))
}

# ---------------------------------------------------------------------------
# Reading the ribbon XML
# ---------------------------------------------------------------------------

# A control whose label is fetched at load carries getLabel. One whose label is
# written into the XML does not, and is not a tag.
ribbon_tags <- function(paths) {
  paths <- paths[file.exists(paths)]
  if (!length(paths)) {
    return(NULL)
  }
  rows <- lapply(paths, function(path) {
    text <- paste(readLines(path, warn = FALSE), collapse = "\n")
    elements <- unlist(regmatches(text, gregexpr("<[a-zA-Z]+[^>]*>", text)))
    # The XML is hand-written and spaces around `=` come and go, so both
    # `getLabel="x"` and `getLabel = "x"` have to be read.
    fetched <- grepl("getLabel[ ]*=", elements)
    has_id <- grepl("[ ]id[ ]*=[ ]*\"", elements)
    fetched <- fetched[has_id]
    elements <- elements[has_id]
    if (!length(elements)) {
      return(NULL)
    }
    ids <- sub('.*[ ]id[ ]*=[ ]*"([^"]+)".*', "\\1", elements)
    keep <- grepl("^[A-Za-z_][A-Za-z0-9_]*$", ids)
    if (!any(keep)) {
      return(NULL)
    }
    data.frame(
      tag = ids[keep],
      file = path,
      fetched = fetched[keep],
      stringsAsFactors = FALSE
    )
  })
  rows <- Filter(Negate(is.null), rows)
  if (!length(rows)) {
    return(NULL)
  }
  rows <- do.call(rbind, rows)
  rows[!duplicated(rows[c("tag", "fetched")]), , drop = FALSE]
}

ribbon_ids <- function(rows) {
  if (is.null(rows)) {
    return(character())
  }
  unique(rows$tag[rows$fetched])
}

# A control that never asks for its label: it either carries a plain `label=`
# written into the XML, or carries no label at all. A table row keyed on one is
# never read, and what the user sees stays the same in every language.
ribbon_fixed_ids <- function(rows) {
  if (is.null(rows)) {
    return(character())
  }
  setdiff(unique(rows$tag[!rows$fetched]), ribbon_ids(rows))
}

LINELIST_RIBBONS <- c(
  "ribbons/_ribbontemplate_main/ribbon.xml",
  "ribbons/_ribbontemplate_dev/ribbon.xml"
)
DESIGNER_RIBBONS <- c(
  "ribbons/designer/ribbon.xml",
  "ribbons/designer_mock/ribbon.xml"
)

# ---------------------------------------------------------------------------
# Reading the designer binary
# ---------------------------------------------------------------------------

# An .xlsb part stores its text as UTF-16LE, and a record may start on an odd
# byte, so both alignments are read at once: every byte that is an identifier
# character AND is followed by a zero byte is kept, then the kept positions are
# cut into runs two bytes apart. That is what a UTF-16LE identifier looks like
# whichever byte it starts on.
utf16_identifiers <- function(path) {
  bytes <- readBin(path, "raw", file.info(path)$size)
  values <- as.integer(bytes)
  total <- length(values)
  if (total < 4L) {
    return(character())
  }

  idx <- seq_len(total - 1L)
  ch <- values[idx]
  keep <- values[idx + 1L] == 0L &
    ((ch >= 48L & ch <= 57L) |
      (ch >= 65L & ch <= 90L) |
      (ch >= 97L & ch <= 122L) |
      ch == 95L)

  positions <- idx[keep]
  if (!length(positions)) {
    return(character())
  }

  breaks <- c(TRUE, diff(positions) != 2L)
  groups <- cumsum(breaks)
  letters_found <- rawToChar(as.raw(values[positions]), multiple = TRUE)
  words <- vapply(
    split(letters_found, groups),
    function(part) paste(part, collapse = ""),
    character(1),
    USE.NAMES = FALSE
  )
  unique(words[nchar(words) > 2L])
}

# The designer sheet the translation runs on. TranslateDesigner is handed
# ThisWorkbook.Worksheets("Main") and nothing else, so a button on the Dev sheet
# is never translated and is not a row the table owes.
MAIN_SHEET <- "Main"

# workbook.bin lists its sheets as a relationship id followed straight away by
# the sheet name, each written as a four-byte character count and then UTF-16LE
# text. Reading the pair back gives the rId of the Main sheet, which the
# package relationships then turn into a worksheet part.
main_sheet_part <- function(work_dir) {
  book <- file.path(work_dir, "xl", "workbook.bin")
  if (!file.exists(book)) {
    return(NA_character_)
  }
  bytes <- readBin(book, "raw", file.info(book)$size)
  values <- as.integer(bytes)

  wanted <- utf8ToInt(MAIN_SHEET)
  target <- as.integer(rbind(wanted, 0L))
  span <- length(target)

  rel_id <- NA_character_
  for (start in seq_len(length(values) - span - 8L)) {
    if (!identical(values[start:(start + span - 1L)], target)) {
      next
    }
    # The count sitting in the four bytes ahead of the text must be its length,
    # or this is the name turning up inside some other string.
    count <- values[start - 4L] + 256L * values[start - 3L]
    if (count != nchar(MAIN_SHEET)) {
      next
    }
    # The relationship id is the string that ends where this count begins.
    tail_end <- start - 5L
    letters_back <- integer(0)
    walk <- tail_end
    while (
      walk > 1L && values[walk] == 0L && is_identifier_byte(values[walk - 1L])
    ) {
      letters_back <- c(values[walk - 1L], letters_back)
      walk <- walk - 2L
    }
    if (length(letters_back)) {
      rel_id <- rawToChar(as.raw(letters_back))
    }
    break
  }

  if (is.na(rel_id) || !grepl("^rId[0-9]+$", rel_id)) {
    return(NA_character_)
  }

  rels <- file.path(work_dir, "xl", "_rels", "workbook.bin.rels")
  if (!file.exists(rels)) {
    return(NA_character_)
  }
  text <- paste(readLines(rels, warn = FALSE), collapse = "")
  entries <- unlist(regmatches(text, gregexpr("<Relationship [^>]*>", text)))
  hit <- entries[grepl(paste0('Id="', rel_id, '"'), entries, fixed = TRUE)]
  if (!length(hit)) {
    return(NA_character_)
  }
  sub('.*Target="([^"]+)".*', "\\1", hit[1])
}

# Only the shapes on the Main sheet, and only those carrying text. A picture
# has no caption to translate, and TranslateShapes writes into TextFrame.
main_sheet_shapes <- function(work_dir) {
  part <- main_sheet_part(work_dir)
  if (is.na(part)) {
    return(character())
  }
  rels <- file.path(
    work_dir,
    "xl",
    "worksheets",
    "_rels",
    paste0(basename(part), ".rels")
  )
  if (!file.exists(rels)) {
    return(character())
  }
  text <- paste(readLines(rels, warn = FALSE), collapse = "")
  targets <- unlist(regmatches(text, gregexpr('Target="[^"]+"', text)))
  targets <- gsub('^Target="|"$', "", targets)
  drawings <- targets[grepl("drawings/drawing[0-9]+\\.xml$", targets)]
  if (!length(drawings)) {
    return(character())
  }

  unlist(lapply(drawings, function(one) {
    path <- file.path(work_dir, "xl", sub("^\\.\\./", "", one))
    if (!file.exists(path)) {
      return(character())
    }
    body <- paste(readLines(path, warn = FALSE), collapse = "\n")
    # One block per shape, cut at the anchor starts. A backreference to close
    # the right tag is not worth the risk here: one anchor type nested in
    # another swallows the whole drawing into a single match.
    starts <- unlist(gregexpr("<xdr:(twoCell|oneCell|absolute)Anchor", body))
    starts <- starts[starts > 0]
    if (!length(starts)) {
      return(character())
    }
    ends <- c(starts[-1] - 1L, nchar(body))
    anchors <- substring(body, starts, ends)

    named <- vapply(
      anchors,
      function(block) {
        if (!grepl("<a:t>", block, fixed = TRUE)) {
          return(NA_character_)
        }
        hit <- regmatches(block, regexpr('name="[^"]+"', block))
        if (!length(hit)) {
          return(NA_character_)
        }
        gsub('^name="|"$', "", hit)
      },
      character(1),
      USE.NAMES = FALSE
    )
    named[!is.na(named)]
  }))
}

read_designer_binary <- function(path) {
  if (!file.exists(path)) {
    return(NULL)
  }
  work_dir <- file.path(tempdir(), "translation-coverage-xlsb")
  unlink(work_dir, recursive = TRUE)
  dir.create(work_dir, recursive = TRUE, showWarnings = FALSE)
  extracted <- tryCatch(
    utils::unzip(path, exdir = work_dir),
    error = function(err) character()
  )
  if (!length(extracted)) {
    return(NULL)
  }

  book <- file.path(work_dir, "xl", "workbook.bin")
  defined_names <- if (file.exists(book)) utf16_identifiers(book) else
    character()

  shapes <- main_sheet_shapes(work_dir)

  unlink(work_dir, recursive = TRUE)
  list(names = unique(defined_names), shapes = unique(shapes))
}

# ---------------------------------------------------------------------------
# Reading the translation workbook
# ---------------------------------------------------------------------------

LANGUAGES <- c("ENG", "FRA", "SPA", "POR", "ARA")

read_translation_tables <- function(path) {
  book <- loadWorkbook(path)
  sheets <- getSheetNames(path)

  out <- list()
  for (sheet in sheets) {
    found <- getTables(book, sheet)
    if (!length(found)) {
      next
    }
    for (position in seq_along(found)) {
      table_name <- tolower(as.character(found[position]))
      address <- paste0(sheet, "!", names(found)[position])
      values <- suppressMessages(read_excel(path, range = address))
      values <- as.data.frame(values, stringsAsFactors = FALSE)
      names(values)[1] <- "tag"
      values$tag <- as.character(values$tag)
      values <- values[
        !is.na(values$tag) & nzchar(trimws(values$tag)),
        ,
        drop = FALSE
      ]
      out[[table_name]] <- list(
        sheet = sheet,
        address = address,
        rows = values
      )
    }
  }
  out
}

tags_of <- function(tables, table_name) {
  entry <- tables[[table_name]]
  if (is.null(entry)) {
    return(character())
  }
  trimws(entry$rows$tag)
}

# ---------------------------------------------------------------------------
# Gathering
# ---------------------------------------------------------------------------

say("Reading ", workbook_path)
tables <- read_translation_tables(workbook_path)

expected_tables <- c(
  "t_tradllshapes",
  "t_tradllmsg",
  "t_tradllforms",
  "t_tradllribbon",
  "t_tradmsg",
  "t_tradrange",
  "t_tradshape",
  "t_traddrop"
)
absent_tables <- setdiff(expected_tables, names(tables))

say("Reading src/classes and src/modules")
source_tags <- read_source_tags()

say("Reading the form binaries")
form_read <- read_form_tags()

say("Reading the ribbon XML")
linelist_ribbon <- ribbon_tags(LINELIST_RIBBONS)
designer_ribbon <- ribbon_tags(DESIGNER_RIBBONS)

say("Reading src/bin/designer/designer.xlsb")
designer_binary <- read_designer_binary("src/bin/designer/designer.xlsb")

all_table_tags <- unique(unlist(lapply(names(tables), function(one) {
  tags_of(tables, one)
})))

# ---------------------------------------------------------------------------
# The message tables: one pool of literals, two tables that may carry them
# ---------------------------------------------------------------------------

message_tables <- c(
  "t_tradllmsg",
  "t_tradmsg",
  "t_tradllshapes",
  "t_tradshape",
  "t_tradrange",
  "t_traddrop"
)
message_known <- unique(unlist(lapply(message_tables, function(one) {
  tags_of(tables, one)
})))

product_tags <- source_tags[source_tags$side != "setup", , drop = FALSE]
missing_messages <- unique(product_tags$tag[
  !product_tags$tag %in% all_table_tags
])
missing_messages <- sort(missing_messages)

message_use <- function(tag) {
  rows <- product_tags[product_tags$tag == tag, , drop = FALSE]
  sides <- unique(rows$side)
  guess <- if (length(sides) == 1L) sides else paste(sides, collapse = " + ")
  list(
    guess = guess,
    files = unique(rows$file)
  )
}

# A message row nobody asks for. Only the literal readers can answer this, so
# the ribbon ids that share t_tradmsg are taken out of the question first.
designer_ribbon_ids <- ribbon_ids(designer_ribbon)

dead_llmsg <- setdiff(tags_of(tables, "t_tradllmsg"), source_tags$tag)
dead_desmsg <- setdiff(
  tags_of(tables, "t_tradmsg"),
  c(source_tags$tag, designer_ribbon_ids)
)
dead_llshapes <- setdiff(tags_of(tables, "t_tradllshapes"), source_tags$tag)


missing_designer_ribbon <- setdiff(
  designer_ribbon_ids,
  tags_of(tables, "t_tradmsg")
)

# ---------------------------------------------------------------------------
# The form table
# ---------------------------------------------------------------------------

if (is.null(form_read)) {
  form_tags_used <- character()
  form_rows <- NULL
  unnamed_controls <- NULL
} else {
  form_rows <- form_read$rows
  form_tags_used <- unique(form_rows$tag[!form_rows$unnamed])
  unnamed_controls <- form_rows[form_rows$unnamed, , drop = FALSE]
}

missing_forms <- sort(setdiff(form_tags_used, tags_of(tables, "t_tradllforms")))
dead_forms <- if (is.null(form_read)) {
  character()
} else {
  sort(setdiff(tags_of(tables, "t_tradllforms"), form_tags_used))
}

form_of <- function(tag) {
  if (is.null(form_rows)) {
    return("")
  }
  paste(unique(form_rows$form[form_rows$tag == tag]), collapse = ", ")
}

# ---------------------------------------------------------------------------
# The ribbon tables
# ---------------------------------------------------------------------------

linelist_ribbon_ids <- ribbon_ids(linelist_ribbon)

# Rows the table carries for a control the XML labels itself.
fixed_label_rows <- sort(c(
  intersect(
    ribbon_fixed_ids(linelist_ribbon),
    tags_of(tables, "t_tradllribbon")
  ),
  intersect(ribbon_fixed_ids(designer_ribbon), tags_of(tables, "t_tradmsg"))
))

missing_linelist_ribbon <- sort(setdiff(
  linelist_ribbon_ids,
  tags_of(tables, "t_tradllribbon")
))
dead_linelist_ribbon <- if (is.null(linelist_ribbon)) {
  character()
} else {
  sort(setdiff(
    tags_of(tables, "t_tradllribbon"),
    c(linelist_ribbon_ids, fixed_label_rows)
  ))
}
# ---------------------------------------------------------------------------
# The two tables keyed on the designer file
# ---------------------------------------------------------------------------

if (is.null(designer_binary)) {
  candidate_ranges <- character()
  dead_ranges <- character()
  missing_shapes <- character()
  dead_shapes <- character()
  dead_drop <- character()
} else {
  # The RNG_Lab prefix is what the labels are named, so a name outside it is a
  # range the code reads rather than a label the user sees. Which SHEET a name
  # points at cannot be read out of workbook.bin without walking the parsed
  # formula, and TranslateRanges only ever runs on Main, so these come back as
  # candidates to check rather than as rows the table owes.
  label_names <- grep("^RNG_Lab", designer_binary$names, value = TRUE)
  candidate_ranges <- sort(setdiff(label_names, tags_of(tables, "t_tradrange")))
  dead_ranges <- sort(setdiff(
    tags_of(tables, "t_tradrange"),
    c(designer_binary$names, source_tags$tag)
  ))

  shape_names <- grep("^SHP_", designer_binary$shapes, value = TRUE)
  missing_shapes <- sort(setdiff(shape_names, tags_of(tables, "t_tradshape")))
  dead_shapes <- sort(setdiff(
    tags_of(tables, "t_tradshape"),
    c(designer_binary$shapes, source_tags$tag)
  ))

  # TranslateDropdowns resolves each row through mainsh.Range(tag), so a
  # dropdown row is keyed on a defined name exactly as a label row is.
  dead_drop <- sort(setdiff(
    tags_of(tables, "t_traddrop"),
    c(designer_binary$names, source_tags$tag)
  ))
}

# ---------------------------------------------------------------------------
# Blank and untranslated cells
# ---------------------------------------------------------------------------

blank_cells <- function() {
  rows <- list()
  for (table_name in names(tables)) {
    values <- tables[[table_name]]$rows
    present <- intersect(LANGUAGES, names(values))
    if (!length(present)) {
      next
    }
    for (language in present) {
      column <- values[[language]]
      empty <- is.na(column) | !nzchar(trimws(as.character(column)))
      if (!any(empty)) {
        next
      }
      rows[[length(rows) + 1L]] <- data.frame(
        table = table_name,
        tag = values$tag[empty],
        language = language,
        stringsAsFactors = FALSE
      )
    }
  }
  if (!length(rows)) {
    return(NULL)
  }
  do.call(rbind, rows)
}

# A row whose five languages all read the same is a row somebody added and never
# translated. Codes and names are meant to be the same everywhere, so a value
# that is a bare identifier is left alone.
untranslated_rows <- function() {
  rows <- list()
  for (table_name in names(tables)) {
    values <- tables[[table_name]]$rows
    present <- intersect(LANGUAGES, names(values))
    if (length(present) < 2L) {
      next
    }
    texts <- as.data.frame(
      lapply(values[present], function(one) trimws(as.character(one))),
      stringsAsFactors = FALSE
    )
    same <- apply(texts, 1, function(one) {
      one <- one[!is.na(one) & nzchar(one)]
      length(one) == length(present) && length(unique(one)) == 1L
    })
    looks_like_a_word <- !grepl("^[A-Za-z_][A-Za-z0-9_]*$", texts[[1]])
    flagged <- same & looks_like_a_word
    if (!any(flagged)) {
      next
    }
    rows[[length(rows) + 1L]] <- data.frame(
      table = table_name,
      tag = values$tag[flagged],
      value = texts[[1]][flagged],
      stringsAsFactors = FALSE
    )
  }
  if (!length(rows)) {
    return(NULL)
  }
  do.call(rbind, rows)
}

blanks <- blank_cells()
untranslated <- untranslated_rows()

duplicate_tags <- function() {
  rows <- list()
  for (table_name in names(tables)) {
    values <- tables[[table_name]]$rows
    repeated <- unique(values$tag[duplicated(trimws(values$tag))])
    if (!length(repeated)) {
      next
    }
    rows[[length(rows) + 1L]] <- data.frame(
      table = table_name,
      tag = repeated,
      stringsAsFactors = FALSE
    )
  }
  if (!length(rows)) {
    return(NULL)
  }
  do.call(rbind, rows)
}

duplicates <- duplicate_tags()

# ---------------------------------------------------------------------------
# The report
# ---------------------------------------------------------------------------

lines <- character()
add <- function(...) {
  lines <<- c(lines, paste0(...))
}

# Every tag list in the report goes through here, so an empty one always reads
# the same. paste0 recycles a zero-length vector to "" when its other arguments
# are longer, so the backticks are added here rather than at the call site.
bullet_list <- function(items, note = NULL) {
  items <- items[!is.na(items) & nzchar(items)]
  if (!length(items)) {
    add("_Nothing._")
    add("")
    return(invisible(NULL))
  }
  if (!is.null(note)) {
    add(note)
    add("")
  }
  for (item in items) {
    add("- `", item, "`")
  }
  add("")
}

table_row <- function(table_name, missing_count, dead_count, note) {
  entry <- tables[[table_name]]
  total <- if (is.null(entry)) 0L else nrow(entry$rows)
  add(
    "| ",
    table_name,
    " | ",
    total,
    " | ",
    missing_count,
    " | ",
    dead_count,
    " | ",
    note,
    " |"
  )
}

# The four tables keyed on string literals share one pool of missing tags.
POOLED <- "see the message list"

add("# Translation coverage")
add("")
add(
  "Written by `scripts/devtools/translation-coverage.R` on ",
  format(Sys.time(), "%Y-%m-%d %H:%M"),
  "."
)
add("")
add("Workbook read: `", workbook_path, "`.")
add("")
add(
  "**Missing** means the product asks for the tag and no table row carries ",
  "it. The screen then shows the tag itself. **Dead** means a row nobody ",
  "asks for."
)
add("")

add("## Summary")
add("")
add("| table | rows | missing | dead | oracle |")
add("| --- | --- | --- | --- | --- |")
table_row(
  "t_tradllshapes",
  POOLED,
  length(dead_llshapes),
  "VBA string literals"
)
table_row("t_tradllmsg", POOLED, length(dead_llmsg), "VBA string literals")
table_row(
  "t_tradllforms",
  length(missing_forms),
  length(dead_forms),
  if (is.null(form_read)) "not checked, no form folder on disk" else
    paste0("control names in `", form_read$folder, "`")
)
table_row(
  "t_tradllribbon",
  length(missing_linelist_ribbon),
  length(dead_linelist_ribbon),
  "getLabel ids in the linelist ribbon templates"
)
table_row(
  "t_tradmsg",
  length(missing_designer_ribbon),
  length(dead_desmsg),
  "VBA literals + getLabel ids in the designer ribbon"
)
table_row(
  "t_tradrange",
  paste0(length(candidate_ranges), " to check"),
  length(dead_ranges),
  if (is.null(designer_binary)) "not checked, designer.xlsb absent" else
    "defined names in designer.xlsb"
)
table_row(
  "t_tradshape",
  length(missing_shapes),
  length(dead_shapes),
  if (is.null(designer_binary)) "not checked, designer.xlsb absent" else
    "shape names in designer.xlsb"
)
table_row(
  "t_traddrop",
  0L,
  length(dead_drop),
  if (is.null(designer_binary)) "not checked, designer.xlsb absent" else
    "defined names in designer.xlsb"
)
add("")
add(
  "Message tags the code asks for and no table carries: **",
  length(missing_messages),
  "**. They are listed on their own below, ",
  "because a string literal does not say which of the two sheets should ",
  "carry it."
)
add("")

if (length(absent_tables)) {
  add(
    "Tables the workbook does not hold: ",
    paste0("`", absent_tables, "`", collapse = ", "),
    "."
  )
  add("")
}

# --- missing ---------------------------------------------------------------

add("## Missing: message tags")
add("")
if (!length(missing_messages)) {
  add("_Nothing._")
  add("")
} else {
  add(
    "The guess is read off the folder the file sits in. The file list ",
    "under it is what the guess is made from."
  )
  add("")
  add("| tag | sheet to fill | used in |")
  add("| --- | --- | --- |")
  for (tag in missing_messages) {
    use <- message_use(tag)
    sheet <- switch(
      use$guess,
      designer = "DesignerTranslation / t_tradmsg",
      linelist = "LinelistTranslation / t_tradllmsg",
      paste0("either (", use$guess, ")")
    )
    add(
      "| `",
      tag,
      "` | ",
      sheet,
      " | ",
      paste0("`", basename(use$files), "`", collapse = ", "),
      " |"
    )
  }
  add("")
}

add("## Missing: form controls")
add("")
if (is.null(form_read)) {
  add(
    "Not checked. No form folder on disk. Export the forms into ",
    "`.mock/forms/designer`, or run the headless build once so ",
    "`.test-runner/forms/merged` is filled, then run this again."
  )
  add("")
} else if (!length(missing_forms)) {
  add("_Nothing._")
  add("")
} else {
  add("| tag | form |")
  add("| --- | --- |")
  for (tag in missing_forms) {
    add("| `", tag, "` | ", form_of(tag), " |")
  }
  add("")
}

add("## Missing: linelist ribbon")
add("")
bullet_list(missing_linelist_ribbon)

add("## Missing: designer ribbon")
add("")
bullet_list(missing_designer_ribbon)

add("## Candidates: designer worksheet labels")
add("")
if (is.null(designer_binary)) {
  add(
    "Not checked. `src/bin/designer/designer.xlsb` is absent. Pull it with ",
    "`scripts/release/pull-assets.sh` and run this again."
  )
  add("")
} else {
  bullet_list(
    candidate_ranges,
    note = paste0(
      "A defined name spelled like a label that the table does not carry. ",
      "The file does not say which sheet a name points at, and only Main is ",
      "translated, so open the workbook before adding a row: a name on an ",
      "internal sheet such as `__pass` is not a label anybody reads."
    )
  )
}

add("## Missing: designer worksheet buttons")
add("")
if (is.null(designer_binary)) {
  add("Not checked, same reason.")
  add("")
} else {
  bullet_list(missing_shapes)
}

# --- dead ------------------------------------------------------------------

add("## Dead rows: linelist messages (t_tradllmsg)")
add("")
bullet_list(dead_llmsg)

add("## Dead rows: linelist shapes (t_tradllshapes)")
add("")
bullet_list(dead_llshapes)

add("## Dead rows: linelist forms (t_tradllforms)")
add("")
if (is.null(form_read)) {
  add("Not checked. No form folder on disk.")
  add("")
} else {
  bullet_list(dead_forms)
}

add("## Dead rows: linelist ribbon (t_tradllribbon)")
add("")
bullet_list(dead_linelist_ribbon)

add("## Dead rows: designer messages and ribbon (t_tradmsg)")
add("")
bullet_list(
  dead_desmsg,
  note = paste0(
    "The designer workbook's own document modules live inside ",
    "`designer.xlsb` and nowhere in `src`, so a tag one of them asks for on ",
    "its own reads as dead here. Check the file before deleting a row."
  )
)

add("## Dead rows: designer worksheet labels (t_tradrange)")
add("")
if (is.null(designer_binary)) {
  add("Not checked. `src/bin/designer/designer.xlsb` is absent.")
  add("")
} else {
  bullet_list(dead_ranges)
}

add("## Dead rows: designer worksheet buttons (t_tradshape)")
add("")
if (is.null(designer_binary)) {
  add("Not checked, same reason.")
  add("")
} else {
  bullet_list(dead_shapes)
}

add("## Dead rows: designer dropdowns (t_traddrop)")
add("")
if (is.null(designer_binary)) {
  add("Not checked. `src/bin/designer/designer.xlsb` is absent.")
  add("")
} else {
  bullet_list(dead_drop)
}

# --- cells to fill ---------------------------------------------------------

add("## Cells left blank")
add("")
if (is.null(blanks)) {
  add("_Nothing._")
  add("")
} else {
  add("| table | tag | language |")
  add("| --- | --- | --- |")
  for (idx in seq_len(nrow(blanks))) {
    add(
      "| ",
      blanks$table[idx],
      " | `",
      blanks$tag[idx],
      "` | ",
      blanks$language[idx],
      " |"
    )
  }
  add("")
}

add("## Rows that read the same in all five languages")
add("")
if (is.null(untranslated)) {
  add("_Nothing._")
  add("")
} else {
  add(
    "A row added in one language and copied across. Some are right on ",
    "purpose -- a product name, a unit."
  )
  add("")
  add("| table | tag | value |")
  add("| --- | --- | --- |")
  for (idx in seq_len(nrow(untranslated))) {
    add(
      "| ",
      untranslated$table[idx],
      " | `",
      untranslated$tag[idx],
      "` | ",
      untranslated$value[idx],
      " |"
    )
  }
  add("")
}

add("## Rows for a ribbon control that never asks for its label")
add("")
if (!length(fixed_label_rows)) {
  add("_Nothing._")
  add("")
} else {
  bullet_list(
    fixed_label_rows,
    note = paste0(
      "The control is in the ribbon XML, but it carries a plain `label=` or ",
      "no label at all rather than `getLabel=`. The row is never read and ",
      "what the user sees stays the same in every language. Either drop the ",
      "row or give the control a getLabel callback."
    )
  )
}

add("## Tags carried twice in one table")
add("")
if (is.null(duplicates)) {
  add("_Nothing._")
  add("")
} else {
  for (idx in seq_len(nrow(duplicates))) {
    add("- `", duplicates$tag[idx], "` in ", duplicates$table[idx])
  }
  add("")
}

add("## Controls left with their default name")
add("")
if (is.null(unnamed_controls) || !nrow(unnamed_controls)) {
  add("_Nothing._")
  add("")
} else {
  add(
    "A control still called `Label1` cannot be translated: no table row ",
    "will ever be keyed on that name. Rename it on the form first."
  )
  add("")
  for (idx in seq_len(nrow(unnamed_controls))) {
    add("- `", unnamed_controls$tag[idx], "` on ", unnamed_controls$form[idx])
  }
  add("")
}

# --- what was read ---------------------------------------------------------

add("## What was read")
add("")
add(
  "- VBA source: `src/classes` and `src/modules`, ",
  length(unique(source_tags$file)),
  " files naming a tag."
)
add(
  "- Forms: ",
  if (is.null(form_read)) "none on disk." else
    paste0(
      "`",
      form_read$folder,
      "`, ",
      length(unique(form_rows$form)),
      " forms naming a tag."
    )
)
add(
  "- Ribbons: ",
  paste0(
    "`",
    c(LINELIST_RIBBONS, DESIGNER_RIBBONS)[
      file.exists(c(LINELIST_RIBBONS, DESIGNER_RIBBONS))
    ],
    "`",
    collapse = ", "
  ),
  "."
)
add(
  "- Designer binary: ",
  if (is.null(designer_binary)) "absent." else
    paste0(
      "`src/bin/designer/designer.xlsb`, ",
      length(designer_binary$names),
      " defined names, ",
      length(designer_binary$shapes),
      " shapes."
    )
)
add("")
add(
  "Left out on purpose: `src/tests`, `src/classes/stale`, and the setup and ",
  "master-setup folders. The setup workbooks carry their own translation ",
  "table and are not in this workbook."
)
add("")

dir.create(dirname(out_path), recursive = TRUE, showWarnings = FALSE)
writeLines(lines, out_path)

finding_count <- length(missing_messages) +
  length(missing_forms) +
  length(missing_linelist_ribbon) +
  length(missing_designer_ribbon) +
  length(missing_shapes) +
  length(dead_llmsg) +
  length(dead_llshapes) +
  length(dead_forms) +
  length(dead_linelist_ribbon) +
  length(dead_desmsg) +
  length(dead_ranges) +
  length(dead_shapes) +
  length(dead_drop)

say("Wrote ", out_path)
say("designer label names to check: ", length(candidate_ranges))
say(
  "missing ",
  length(missing_messages) +
    length(missing_forms) +
    length(missing_linelist_ribbon) +
    length(missing_designer_ribbon) +
    length(missing_shapes),
  ", dead ",
  length(dead_llmsg) +
    length(dead_llshapes) +
    length(dead_forms) +
    length(dead_linelist_ribbon) +
    length(dead_desmsg) +
    length(dead_ranges) +
    length(dead_shapes) +
    length(dead_drop)
)

quit(status = if (finding_count > 0L) 1L else 0L)

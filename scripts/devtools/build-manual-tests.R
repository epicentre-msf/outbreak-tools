#!/usr/bin/env Rscript
#
# build-manual-tests.R -- build the manual test workbook from its yaml.
#
#   Rscript scripts/devtools/build-manual-tests.R [input.yml] [output.xlsx]
#
# Defaults:
#   input   src/tests/manual-tests.yml
#   output  src/tests/.input/manual_tests_list.xlsx
#
# THE YAML IS THE SOURCE. The workbook is overwritten whole on every run, so a
# test line typed into Excel is lost at the next build. Add the line to the
# yaml instead. The one thing the workbook holds and the yaml does not is the
# Response column, which is why the build always leaves it empty: fill in a
# copy, and keep the copy.
#
# WHAT IT WRITES
# -----------------------------------------------------------------------------
# One worksheet per `workbooks:` entry, in the order they are written. Each
# sheet carries two ListObjects side by side:
#
#      B            C                D            G      H      I
#   +--------+--------------------+----------+  +------+------+----------+
#   | Area   | Test that:         | Response |  | Area | Test | Response |
#   +--------+--------------------+----------+  +------+------+----------+
#     <sheet>_functionalities                     <sheet>_features
#
# A table name has to be unique across the whole workbook, so each one carries
# its sheet: `setups_functionalities`, `designer_features`.
#
# A section with no tests still gets its table, one blank row wide, so the
# dropdown and the colours are there to type into.
#
# THE COLOURS
# -----------------------------------------------------------------------------
# Response is a dropdown of Yes,No that allows a blank, and two conditional
# rules paint the whole row from it:
#
#   (empty)  not run yet   no colour
#   Yes      passed        green  (font 006100 on fill C6EFCE)
#   No       FAILED        red    (font 9C0006 on fill FFC7CE)
#
# The rules are `expression` rules anchored on the Response column of the row
# ($D4, $I4), so the colour follows the row when the table is sorted or
# filtered. Leaving the blank uncoloured is the point: a red row is a failure
# a person saw, never a test nobody has reached yet.
#
# WHY openxlsx2 AND NOT openxlsx
# -----------------------------------------------------------------------------
# openxlsx writes a list validation into the x14 extension block. Excel reads
# it, WPS Office in the field often does not. openxlsx2 writes the plain
# <dataValidations> element, which is what the hand-made workbook carried.
#
# Both libraries need a hand on the conditional fill; see the note beside the
# dxfs styles below.
#
# WHAT IT CHECKS BEFORE WRITING
# -----------------------------------------------------------------------------
# A missing key or an empty test line stops the run. A sentence repeated inside
# one table is printed as a note and does not stop it -- the list carries a few
# near-repeats on purpose, so that stays a person's decision.
# =============================================================================

suppressPackageStartupMessages({
  library(openxlsx2)
  library(yaml)
})

args <- commandArgs(trailingOnly = TRUE)
input <- if (length(args) >= 1L) args[1] else "src/tests/manual-tests.yml"
output <- if (length(args) >= 2L) args[2] else "src/tests/.input/manual_tests_list.xlsx"

if (!file.exists(input)) stop("no yaml at ", input, call. = FALSE)

# --- the look, all of it -----------------------------------------------------
KINDS <- c("functionalities", "features")
TITLES <- c(functionalities = "Test of functionalities", features = "Test of features")
START_COL <- c(functionalities = 2L, features = 7L) # B and G
TITLE_ROW <- 2L
HEADER_ROW <- 3L
HEADERS <- c("Area", "Test that:", "Response")
WIDTHS <- c(20, 62, 16) # Area, Test that:, Response
GAP_WIDTH <- 4 # the two columns between the tables
LINE_HEIGHT <- 17 # one line of text, in points
WRAP_CHARS <- 58 # characters that fit on one line of the test column
TABLE_STYLE <- "TableStyleLight1"

spec <- yaml::read_yaml(input)

response <- spec$response
if (is.null(response$values)) stop("the yaml has no `response: values:`", call. = FALSE)
values <- unlist(response$values)
pass_value <- if (is.null(response$pass)) values[1] else as.character(response$pass)
fail_value <- if (is.null(response$fail)) values[2] else as.character(response$fail)
validation <- paste0('"', paste(values, collapse = ","), '"')

if (is.null(spec$workbooks) || !length(spec$workbooks)) {
  stop("the yaml has no `workbooks:`", call. = FALSE)
}

# Flatten one kind of one workbook into the rows of its table. A test is either
# the sentence on its own or a `test:`/`note:` pair; both land in the same three
# columns, and the note travels beside them for the cell comment.
flatten_kind <- function(areas, where) {
  rows <- list()
  for (block in areas) {
    area <- block$area
    if (is.null(area) || !nzchar(trimws(as.character(area)))) {
      stop(where, ": an area block has no `area:` name", call. = FALSE)
    }
    for (entry in block$tests) {
      if (is.list(entry)) {
        text <- entry$test
        note <- if (is.null(entry$note)) NA_character_ else as.character(entry$note)
        # A test line holding a colon and a space is read by YAML as a key, so
        # it arrives here as a list with no `test`. Say that, rather than
        # "the line is empty", which is what it looks like from in here.
        if (is.null(text)) {
          stop(
            where, ", area ", area, ": this line has no `test:` --\n  ",
            paste(names(entry), collapse = ", "),
            "\n  A colon and a space make YAML read the line as a key. ",
            "Wrap the whole sentence in double quotes.",
            call. = FALSE
          )
        }
      } else {
        text <- entry
        note <- NA_character_
      }
      if (!nzchar(trimws(as.character(text)))) {
        stop(where, ", area ", area, ": a test line is empty", call. = FALSE)
      }
      rows[[length(rows) + 1L]] <- list(
        area = as.character(area), test = trimws(as.character(text)), note = note
      )
    }
  }
  data.frame(
    area = vapply(rows, `[[`, "", "area"),
    test = vapply(rows, `[[`, "", "test"),
    note = vapply(rows, function(r) r$note, ""),
    stringsAsFactors = FALSE
  )
}

wb <- wb_workbook()
wb <- wb_add_dxfs_style(wb, "manual_pass",
  font_color = wb_color(hex = "FF006100"), bg_fill = wb_color(hex = "FFC6EFCE")
)
wb <- wb_add_dxfs_style(wb, "manual_fail",
  font_color = wb_color(hex = "FF9C0006"), bg_fill = wb_color(hex = "FFFFC7CE")
)

# openxlsx2 writes the conditional fill as <patternFill patternType="solid">
# with a bgColor and no fgColor. Under a solid pattern the fgColor is the one
# that paints, so a reader that follows the spec to the letter draws nothing.
# Excel writes its own conditional fills with no patternType at all and reads
# the bgColor, which is the form the hand-made workbook carried. Drop the
# attribute so both readings land on the same colour.
wb$styles_mgr$styles$dxfs <- sub(
  '<patternFill patternType="solid">', "<patternFill>",
  wb$styles_mgr$styles$dxfs,
  fixed = TRUE
)

counts <- list()
notes_seen <- character(0)

for (book in spec$workbooks) {
  sheet <- book$sheet
  if (is.null(sheet) || !nzchar(trimws(as.character(sheet)))) {
    stop("a `workbooks:` entry has no `sheet:` name", call. = FALSE)
  }
  sheet <- as.character(sheet)
  wb <- wb_add_worksheet(wb, sheet, grid_lines = FALSE)

  tables <- lapply(KINDS, function(kind) {
    flatten_kind(book[[kind]], paste0(sheet, " / ", kind))
  })
  names(tables) <- KINDS

  # The two tables share their row heights, so measure both before writing.
  n_rows <- max(1L, vapply(tables, nrow, 0L))

  for (kind in KINDS) {
    df <- tables[[kind]]
    col_area <- START_COL[[kind]]
    col_test <- col_area + 1L
    col_resp <- col_area + 2L

    counts[[paste0(sheet, "/", kind)]] <- nrow(df)

    repeated <- unique(df$test[duplicated(df$test)])
    if (length(repeated)) {
      notes_seen <- c(notes_seen, sprintf(
        "%s / %s repeats %d test line(s): %s", sheet, kind, length(repeated),
        paste0('"', repeated, '"', collapse = "; ")
      ))
    }

    # An Excel table needs a data row, so an empty section gets a blank one.
    body <- if (nrow(df)) {
      data.frame(df$area, df$test, NA_character_, stringsAsFactors = FALSE)
    } else {
      data.frame(NA_character_, NA_character_, NA_character_, stringsAsFactors = FALSE)
    }
    names(body) <- HEADERS

    first <- HEADER_ROW + 1L
    last <- HEADER_ROW + nrow(body)

    wb <- wb_add_data(wb, sheet,
      x = TITLES[[kind]],
      dims = wb_dims(rows = TITLE_ROW, cols = col_area), col_names = FALSE
    )
    wb <- wb_add_font(wb, sheet,
      dims = wb_dims(rows = TITLE_ROW, cols = col_area), bold = "1", size = "12"
    )

    wb <- wb_add_data_table(wb, sheet,
      x = body, dims = wb_dims(rows = HEADER_ROW, cols = col_area),
      table_name = paste0(sheet, "_", kind), table_style = TABLE_STYLE,
      with_filter = TRUE, na.strings = NULL
    )

    wb <- wb_add_cell_style(wb, sheet,
      dims = wb_dims(rows = first:last, cols = col_area),
      vertical = "center", wrap_text = "1"
    )
    wb <- wb_add_cell_style(wb, sheet,
      dims = wb_dims(rows = first:last, cols = col_test),
      vertical = "center", horizontal = "left", wrap_text = "1"
    )
    wb <- wb_add_cell_style(wb, sheet,
      dims = wb_dims(rows = first:last, cols = col_resp),
      vertical = "center", horizontal = "center"
    )

    wb <- wb_set_col_widths(wb, sheet, cols = col_area:col_resp, widths = WIDTHS)

    wb <- wb_add_data_validation(wb, sheet,
      dims = wb_dims(rows = first:last, cols = col_resp),
      type = "list", value = validation, allow_blank = TRUE
    )

    # Anchored on the Response column, relative row: the colour travels with
    # the row through a sort or a filter.
    anchor <- paste0("$", int2col(col_resp), first)
    band <- wb_dims(rows = first:last, cols = col_area:col_resp)
    wb <- wb_add_conditional_formatting(wb, sheet,
      dims = band, type = "expression",
      rule = sprintf('%s="%s"', anchor, pass_value), style = "manual_pass"
    )
    wb <- wb_add_conditional_formatting(wb, sheet,
      dims = band, type = "expression",
      rule = sprintf('%s="%s"', anchor, fail_value), style = "manual_fail"
    )

    for (i in which(!is.na(df$note))) {
      wb <- wb_add_comment(wb, sheet,
        dims = wb_dims(rows = HEADER_ROW + i, cols = col_test),
        comment = wb_comment(text = df$note[i], author = basename(input), visible = FALSE)
      )
    }
  }

  gap <- (START_COL[["functionalities"]] + 3L):(START_COL[["features"]] - 1L)
  wb <- wb_set_col_widths(wb, sheet, cols = gap, widths = GAP_WIDTH)

  # Each row is tall enough for the longer of the two sentences sitting on it.
  heights <- vapply(seq_len(n_rows), function(i) {
    lines <- vapply(KINDS, function(kind) {
      df <- tables[[kind]]
      if (i > nrow(df)) {
        return(1L)
      }
      max(1L, as.integer(ceiling(nchar(df$test[i]) / WRAP_CHARS)))
    }, 0L)
    LINE_HEIGHT * max(lines)
  }, 0)
  wb <- wb_set_row_heights(wb, sheet,
    rows = (HEADER_ROW + 1L):(HEADER_ROW + n_rows), heights = heights
  )
  wb <- wb_freeze_pane(wb, sheet, first_active_row = HEADER_ROW + 1L)
}

dir.create(dirname(output), showWarnings = FALSE, recursive = TRUE)
wb_save(wb, output, overwrite = TRUE)

for (n in notes_seen) message("note: ", n)
cat("wrote ", output, "\n", sep = "")
for (nm in names(counts)) cat(sprintf("  %-28s %3d tests\n", nm, counts[[nm]]))
cat(sprintf("  %-28s %3d tests\n", "TOTAL", sum(unlist(counts))))

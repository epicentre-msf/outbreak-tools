#!/usr/bin/env Rscript

# Add rows to one table of a translations workbook.
#
# The translations workbooks stack several tables down one worksheet, with a
# title line and a blank line between them. A new message code has to land
# inside the right table, which means every block under it moves down. openxlsx
# cannot insert rows, so this works on the xlsx parts directly: the sheet XML is
# renumbered, the new rows are written as inline strings, and every table
# reference on the sheet is grown or shifted to match.
#
# Usage:
#
#   Rscript scripts/devtools/add-translation-rows.R \
#     --workbook trads/designer_translations.xlsx \
#     --rows trads/new-rows.csv
#
# The rows file is a csv whose first column is the table name and whose
# remaining columns are the row itself, starting with the message code:
#
#   table,msg_id,ENG,FRA,SPA,POR,ARA
#   t_tradllmsg,MSG_Example,Example,Exemple,Ejemplo,Exemplo,مثال
#
# A code the table already carries is skipped and reported. The workbook is
# rewritten in place, and nothing outside the sheets holding the named tables is
# touched.

suppressMessages({
  library(xml2)
  library(zip)
})

XL_NS <- "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# ---- arguments ---------------------------------------------------------------

parse_args <- function(argv) {
  out <- list(workbook = NULL, rows = NULL)
  idx <- 1
  while (idx <= length(argv)) {
    key <- argv[[idx]]
    value <- if (idx < length(argv)) argv[[idx + 1]] else NA_character_
    if (key == "--workbook") out$workbook <- value
    if (key == "--rows") out$rows <- value
    idx <- idx + 2
  }
  if (is.null(out$workbook) || is.null(out$rows)) {
    stop("both --workbook and --rows are required", call. = FALSE)
  }
  out
}

# ---- cell references ---------------------------------------------------------

# "B21" -> 21
row_of <- function(ref) as.integer(sub("^[A-Z]+", "", ref))

# "B21" -> "B"
column_of <- function(ref) sub("[0-9]+$", "", ref)

# "B21:G190" -> list(first = 21, last = 190, left = "B", right = "G")
parse_ref <- function(ref) {
  ends <- strsplit(ref, ":", fixed = TRUE)[[1]]
  list(
    first = row_of(ends[[1]]),
    last = row_of(ends[[length(ends)]]),
    left = column_of(ends[[1]]),
    right = column_of(ends[[length(ends)]])
  )
}

build_ref <- function(bounds) {
  paste0(bounds$left, bounds$first, ":", bounds$right, bounds$last)
}

# The column letters a table spans, left to right.
columns_of <- function(bounds) {
  letters_all <- c(LETTERS, as.vector(outer(LETTERS, LETTERS, paste0)))
  first <- match(bounds$left, letters_all)
  last <- match(bounds$right, letters_all)
  letters_all[first:last]
}

# ---- the sheet -----------------------------------------------------------

# Move every row at or below `from` down by `count`, cell references included.
shift_rows <- function(sheet_doc, from, count) {
  rows <- xml_find_all(sheet_doc, paste0("//d1:sheetData/d1:row"),
                       ns = c(d1 = XL_NS))
  numbers <- as.integer(xml_attr(rows, "r"))
  moving <- which(numbers >= from)

  # Bottom up, so two rows never hold the same number at once.
  for (idx in rev(moving)) {
    node <- rows[[idx]]
    new_number <- numbers[[idx]] + count
    xml_set_attr(node, "r", as.character(new_number))
    for (cell in xml_find_all(node, "d1:c", ns = c(d1 = XL_NS))) {
      xml_set_attr(cell, "r", paste0(column_of(xml_attr(cell, "r")), new_number))
    }
  }
  invisible(NULL)
}

# The style each column of a table body uses, read off the row above the insert
# point. A new row that carries no style prints white on white in some themes.
styles_of_row <- function(sheet_doc, row_number, wanted_columns) {
  node <- xml_find_first(sheet_doc,
                         sprintf("//d1:sheetData/d1:row[@r='%d']", row_number),
                         ns = c(d1 = XL_NS))
  styles <- setNames(rep(NA_character_, length(wanted_columns)), wanted_columns)
  if (inherits(node, "xml_missing")) return(styles)

  for (cell in xml_find_all(node, "d1:c", ns = c(d1 = XL_NS))) {
    letter <- column_of(xml_attr(cell, "r"))
    if (letter %in% wanted_columns) styles[[letter]] <- xml_attr(cell, "s")
  }
  styles
}

# Write one row of text into the sheet, as inline strings.
#
# The row goes in straight after `after_node`, because Excel reads sheetData in
# document order and refuses a sheet whose rows climb out of order. The node is
# built with no namespace prefix of its own and inherits the sheet's default one
# when the file is written, which is the shape every other row already has.
insert_row <- function(after_node, row_number, letters_wanted, values, styles) {
  row_node <- xml_add_sibling(after_node, "row", .where = "after")
  xml_set_attr(row_node, "r", as.character(row_number))
  xml_set_attr(row_node, "spans",
               paste0(match(letters_wanted[[1]], LETTERS), ":",
                      match(letters_wanted[[length(letters_wanted)]], LETTERS)))

  for (position in seq_along(letters_wanted)) {
    letter <- letters_wanted[[position]]
    text <- values[[position]]
    if (is.na(text)) text <- ""

    cell <- xml_add_child(row_node, "c")
    xml_set_attr(cell, "r", paste0(letter, row_number))
    if (!is.na(styles[[letter]])) xml_set_attr(cell, "s", styles[[letter]])

    if (nzchar(text)) {
      xml_set_attr(cell, "t", "inlineStr")
      is_node <- xml_add_child(cell, "is")
      t_node <- xml_add_child(is_node, "t")
      xml_text(t_node) <- text
      xml_set_attr(t_node, "xml:space", "preserve")
    }
  }
  row_node
}

# The row node the new rows are hung off: the last row of the table before the
# insert, which shift_rows has already renumbered.
row_node_at <- function(sheet_doc, row_number) {
  xml_find_first(sheet_doc,
                 sprintf("//d1:sheetData/d1:row[@r='%d']", row_number),
                 ns = c(d1 = XL_NS))
}

grow_dimension <- function(sheet_doc, count) {
  node <- xml_find_first(sheet_doc, "//d1:dimension", ns = c(d1 = XL_NS))
  if (inherits(node, "xml_missing")) return(invisible(NULL))

  bounds <- parse_ref(xml_attr(node, "ref"))
  bounds$last <- bounds$last + count
  xml_set_attr(node, "ref", build_ref(bounds))
  invisible(NULL)
}

# ---- the tables --------------------------------------------------------------

# Grow the table the rows land in and shift every table below it. A table above
# the insert point is left exactly as it was.
move_table <- function(path, insert_at, count) {
  doc <- read_xml(path)
  node <- xml_root(doc)
  bounds <- parse_ref(xml_attr(node, "ref"))

  if (bounds$last < insert_at - 1) {
    return(invisible(FALSE))
  }

  if (bounds$first >= insert_at) {
    bounds$first <- bounds$first + count
  }
  bounds$last <- bounds$last + count

  xml_set_attr(node, "ref", build_ref(bounds))

  filter_node <- xml_find_first(node, "d1:autoFilter", ns = c(d1 = XL_NS))
  if (!inherits(filter_node, "xml_missing")) {
    xml_set_attr(filter_node, "ref", build_ref(bounds))
  }

  write_xml(doc, path)
  invisible(TRUE)
}

# ---- the run -----------------------------------------------------------------

main <- function(argv) {
  args <- parse_args(argv)
  workbook_path <- normalizePath(args$workbook, mustWork = TRUE)
  new_rows <- read.csv(args$rows, stringsAsFactors = FALSE,
                       colClasses = "character", check.names = FALSE,
                       encoding = "UTF-8")

  if (!"table" %in% names(new_rows)) {
    stop("the rows file needs a `table` column", call. = FALSE)
  }

  work_dir <- file.path(tempdir(), paste0("obt-trad-", Sys.getpid()))
  unlink(work_dir, recursive = TRUE)
  dir.create(work_dir, recursive = TRUE)
  on.exit(unlink(work_dir, recursive = TRUE), add = TRUE)

  zip::unzip(workbook_path, exdir = work_dir)

  table_paths <- list.files(file.path(work_dir, "xl", "tables"),
                            pattern = "\\.xml$", full.names = TRUE)
  table_names <- vapply(table_paths, function(p) {
    tolower(xml_attr(xml_root(read_xml(p)), "displayName"))
  }, character(1))

  # Which sheet each table belongs to, read off the sheet relationships.
  sheet_files <- list.files(file.path(work_dir, "xl", "worksheets"),
                            pattern = "^sheet[0-9]+\\.xml$", full.names = TRUE)
  table_sheet <- setNames(rep(NA_character_, length(table_paths)), table_paths)
  for (sheet_file in sheet_files) {
    rels_file <- file.path(dirname(sheet_file), "_rels",
                           paste0(basename(sheet_file), ".rels"))
    if (!file.exists(rels_file)) next
    targets <- xml_attr(xml_children(read_xml(rels_file)), "Target")
    for (target in targets) {
      hit <- table_paths[basename(table_paths) == basename(target)]
      if (length(hit)) table_sheet[[hit]] <- sheet_file
    }
  }

  for (wanted in unique(new_rows$table)) {
    block <- new_rows[new_rows$table == wanted, , drop = FALSE]
    value_columns <- setdiff(names(block), "table")

    table_path <- names(table_names)[table_names == tolower(wanted)]
    if (length(table_path) != 1) {
      stop("no table named ", wanted, " in the workbook", call. = FALSE)
    }
    sheet_file <- table_sheet[[table_path]]

    bounds <- parse_ref(xml_attr(xml_root(read_xml(table_path)), "ref"))
    letters_wanted <- columns_of(bounds)
    if (length(letters_wanted) != length(value_columns)) {
      stop("table ", wanted, " has ", length(letters_wanted),
           " columns and the rows file gives ", length(value_columns),
           call. = FALSE)
    }

    insert_at <- bounds$last + 1
    count <- nrow(block)

    sheet_doc <- read_xml(sheet_file)
    styles <- styles_of_row(sheet_doc, bounds$last, letters_wanted)

    shift_rows(sheet_doc, insert_at, count)

    anchor <- row_node_at(sheet_doc, bounds$last)
    if (inherits(anchor, "xml_missing")) {
      stop("table ", wanted, " has no row ", bounds$last, " to hang the new ",
           "rows off", call. = FALSE)
    }

    for (idx in seq_len(count)) {
      anchor <- insert_row(anchor, insert_at + idx - 1, letters_wanted,
                           as.character(block[idx, value_columns]), styles)
    }

    grow_dimension(sheet_doc, count)
    write_xml(sheet_doc, sheet_file)

    same_sheet <- names(table_sheet)[!is.na(table_sheet) &
                                       table_sheet == sheet_file]
    for (other in same_sheet) {
      move_table(other, insert_at, count)
    }

    cat(sprintf("%s: %d row(s) added at row %d\n", wanted, count, insert_at))
  }

  # The content types part goes in first, and every other part keeps the folder
  # it came out of. A flat zip opens as a damaged workbook.
  parts <- list.files(work_dir, recursive = TRUE, all.files = TRUE,
                      no.. = TRUE)
  parts <- c("[Content_Types].xml", setdiff(parts, "[Content_Types].xml"))

  unlink(workbook_path)
  zip::zip(zipfile = workbook_path, files = parts, root = work_dir,
           mode = "mirror")

  cat("written:", workbook_path, "\n")
}

main(commandArgs(trailingOnly = TRUE))

# merge-form-code.R
# =============================================================================
# Put the current FormLogic code into the exported forms, ready to import.
#
#   Rscript scripts/headless/merge-form-code.R [--out <dir>]
#
# A .frm exported from a workbook carries THREE parts, in this order:
#
#   VERSION 5.00                     one line
#   Begin {GUID} F_Name ... End      the control tree, and the .frx beside it
#   Attribute VB_Name = "F_Name"     the module attributes of the form
#   <code>                           the code behind the form
#
# The forms in .mock/forms/designer carry whatever code was in the workbook the
# day they were exported. The code that BELONGS in them lives in
# src/modules/linelistform/FormLogic*.bas, which is what the designer copies
# into each form at deployment (Development.AddFormsCodes). So a form imported
# straight from .mock would run stale logic.
#
# This script rewrites the third part and leaves the first two exactly as they
# are: the header is copied byte for byte through the last Attribute line, then
# the body of the matching FormLogic module is appended with ITS OWN attribute
# lines dropped. A second `Attribute VB_Name` inside a form module is what makes
# an import fail, and the control tree is what the .frx is keyed to, so neither
# is touched.
#
# What it writes: <out>/F_Name.frm (merged) and <out>/F_Name.frx (copied), for
# every form the linelist uses. A form with no FormLogic module keeps its own
# code and is still rewritten. F_Progr is skipped: the designer's progress
# dialog is the one form the transfer never moves.
#
# LINE ENDINGS ARE CRLF, AND THAT IS LOAD-BEARING
# -----------------------------------------------------------------------------
# The VBE reads a .frm as a Windows text file whatever the host. Given LF
# endings it cannot parse the `Begin ... End` control tree, so it imports the
# file as a STANDARD MODULE whose first lines are the header text -- and it
# reports no error at all. Every .frm written here goes through write_frm for
# exactly that reason; only the binary .frx is copied byte for byte.
# =============================================================================

suppressWarnings(suppressMessages({
  library(fs)
  library(cli)
}))

args <- commandArgs(trailingOnly = TRUE)

repo_root <- tryCatch(
  system("git rev-parse --show-toplevel", intern = TRUE),
  error = function(e) getwd()
)

out_dir <- if (any(args == "--out")) {
  args[which(args == "--out") + 1]
} else {
  path(repo_root, ".test-runner", "forms", "merged")
}

forms_dir <- path(repo_root, ".mock", "forms", "designer")
logic_dir <- path(repo_root, "src", "modules", "linelistform")

# Forms this workflow has no use for. F_Progr is the one form
# Linelist.TransferAllCode never moves: it is the designer's own progress
# dialog, it never reaches a linelist, and emitting it only gave the headless
# import a file to trip over. Owner decision, 2026-08-13.
form_skip <- c("F_Progr")

# The form each FormLogic module belongs to. Every pair here is stated by the
# module's own @ModuleDescription; the table is explicit because two of them
# cannot be derived from the file name -- FormLogicShowHide belongs to
# F_ShowHideLL, and FormLogicEpiWeek names no form at all.
form_logic <- c(
  F_Advanced         = "FormLogicAdvanced",
  F_EpiWeek          = "FormLogicEpiWeek",
  F_Export           = "FormLogicExport",
  F_ExportMig        = "FormLogicExportMig",
  F_Geo              = "FormLogicGeo",
  F_ImportRep        = "FormLogicImportRep",
  F_ShowHideLL       = "FormLogicShowHide",
  F_ShowHidePrint    = "FormLogicShowHidePrint",
  F_ShowHideSave     = "FormLogicShowHideSave",
  F_ShowHideSections = "FormLogicShowHideSections",
  F_ShowVarLabels    = "FormLogicShowVarLabels"
)

# --- the two readers ---------------------------------------------------------

#' The header of a .frm: everything up to and including its last attribute line
#'
#' The attribute block sits directly after the `End` of the control tree, and
#' the code starts on the first line after it. Reading it this way rather than
#' by a fixed line count keeps the script honest against a form that carries
#' more or fewer attributes than its siblings.
frm_header <- function(lines) {
  attribute_lines <- grep("^Attribute ", lines)

  if (!length(attribute_lines)) {
    stop("no Attribute line in the form: it is not a .frm export", call. = FALSE)
  }

  last_attribute <- max(attribute_lines)

  # Guard the shape rather than trusting it: the attributes must be one
  # unbroken block, or "everything above the last one" would swallow code.
  first_attribute <- min(attribute_lines)
  if (!identical(attribute_lines, first_attribute:last_attribute)) {
    stop("the Attribute lines of the form are not one block", call. = FALSE)
  }

  lines[seq_len(last_attribute)]
}

#' The body of a FormLogic module: everything below its own attribute lines
#'
#' A standard module exported from the VBE opens with `Attribute VB_Name`.
#' Copying that into a form module gives the form a second VB_Name and the
#' import fails, which is the whole reason this function exists.
logic_body <- function(lines) {
  attribute_lines <- grep("^Attribute ", lines)
  if (!length(attribute_lines)) {
    return(lines)
  }
  lines[-seq_len(max(attribute_lines))]
}

# --- the merge ---------------------------------------------------------------

dir_create(out_dir, recurse = TRUE)

#' Write one .frm with CRLF endings, whatever the source had
#'
#' EVERY .frm this script emits goes through here, including one that keeps
#' its own code. A .frm is read by the VBE as a Windows text file: given LF
#' endings it cannot parse the `Begin ... End` control tree, and it imports the
#' whole file as a STANDARD MODULE whose first lines are the header text. That
#' is what a plain file copy produced for the one form that was not rewritten,
#' and the VBE said nothing about it -- the import "succeeded".
write_frm <- function(lines, out_path) {
  con <- file(out_path, open = "wb")
  on.exit(close(con))
  writeLines(lines, con = con, sep = "\r\n")
}

merged <- 0L
copied <- 0L
skipped <- 0L

frm_files <- dir_ls(forms_dir, glob = "*.frm")

for (frm in frm_files) {
  form_name <- path_ext_remove(path_file(frm))

  if (form_name %in% form_skip) {
    skipped <- skipped + 1L
    cli_alert_info("{form_name}: not used by the linelist, skipped")
    next
  }

  frx <- path(forms_dir, paste0(form_name, ".frx"))

  out_frm <- path(out_dir, paste0(form_name, ".frm"))
  out_frx <- path(out_dir, paste0(form_name, ".frx"))

  # The .frx carries the controls and is keyed to the control tree, so it
  # travels beside the .frm untouched -- and it is BINARY, so it is the one
  # file here that must never be rewritten as text. A form without one is not
  # an error: a form holding no embedded object ships without it.
  if (file_exists(frx)) {
    file_copy(frx, out_frx, overwrite = TRUE)
  }

  # Single bracket, and the membership test before it: `form_logic[[name]]` on a
  # name the vector does not carry raises "subscript out of bounds" rather than
  # answering NULL.
  if (!form_name %in% names(form_logic)) {
    write_frm(readLines(frm, warn = FALSE), out_frm)
    copied <- copied + 1L
    cli_alert_info("{form_name}: no FormLogic module, kept its own code")
    next
  }

  logic_name <- form_logic[[form_name]]
  logic_file <- path(logic_dir, paste0(logic_name, ".bas"))
  if (!file_exists(logic_file)) {
    cli_abort("{form_name}: {logic_name}.bas is missing from {logic_dir}")
  }

  header <- frm_header(readLines(frm, warn = FALSE))
  body <- logic_body(readLines(logic_file, warn = FALSE))

  write_frm(c(header, body), out_frm)

  merged <- merged + 1L
  cli_alert_success("{form_name}: {logic_name} merged ({length(body)} lines)")
}

# A skipped form left behind by an earlier run would still be imported, since
# the VBA side reads whatever .frm files the folder holds.
for (dead_name in form_skip) {
  for (ext in c(".frm", ".frx")) {
    dead_file <- path(out_dir, paste0(dead_name, ext))
    if (file_exists(dead_file)) {
      file_delete(dead_file)
      cli_alert_info("removed stale {dead_name}{ext} from a previous run")
    }
  }
}

cli_rule()
cli_alert_success(
  "{merged} merged, {copied} kept own code, {skipped} skipped -> {out_dir}"
)

# A form named in the mapping but absent from .mock is worth a word: the
# linelist build asks for each of them by name and a missing one ships a
# linelist short of a form.
missing <- setdiff(names(form_logic), path_ext_remove(path_file(frm_files)))
if (length(missing)) {
  cli_alert_warning("no .frm in {forms_dir} for: {paste(missing, collapse = ', ')}")
}

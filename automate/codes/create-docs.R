library(here)
library(fs)
library(stringr)

source(here("automate/codes/class-doc.R"))

classes_folder <- here("src")

parser <- VBADocParser$new(
  folder = classes_folder,
  proj_path = here::here(),
  output_folder = here("src", "docs")
)

# --- Read .docignore patterns ------------------------------------------------

docignore_path <- here("src", ".docignore")
exclude_files <- character()

if (file_exists(docignore_path)) {
  patterns <- readLines(docignore_path, warn = FALSE)
  patterns <- str_trim(patterns)
  patterns <- patterns[nzchar(patterns) & !str_starts(patterns, "#")]

  # Resolve each pattern against the source tree
  all_cls <- dir_ls(classes_folder, recurse = TRUE, regexp = "\\.cls$")
  rel_paths <- path_rel(all_cls, classes_folder)

  for (pat in patterns) {
    # Directory pattern (ends with /): exclude anything under it
    if (str_ends(pat, "/")) {
      dir_pat <- str_remove(pat, "/$")
      matches <- all_cls[str_detect(rel_paths, fixed(dir_pat))]
    } else if (str_detect(pat, "[*?]")) {
      # Glob pattern
      matches <- all_cls[str_detect(
        path_file(all_cls),
        glob2rx(pat, trim.head = TRUE, trim.tail = TRUE)
      )]
    } else {
      # Exact relative path
      matches <- all_cls[str_detect(rel_paths, fixed(pat))]
    }
    exclude_files <- c(exclude_files, matches)
  }
  exclude_files <- unique(exclude_files)

  if (length(exclude_files) > 0) {
    cli::cli_inform(
      "Excluding {length(exclude_files)} file(s) via .docignore"
    )
  }
}

# --- Parse --------------------------------------------------------------------

args <- commandArgs(trailingOnly = TRUE)

if (length(args) == 0) {
  parser$parse(exclude_files = exclude_files)
} else {
  mode <- tolower(args[[1]])
  parse_mode <- NULL
  target <- NULL

  extract_target <- function(flag, provided_args) {
    if (length(provided_args) > 1) {
      return(provided_args[[2]])
    }
    cli::cli_abort(sprintf("Missing class name after '%s'", flag))
  }

  if (startsWith(mode, "--class=")) {
    parse_mode <- "class"
    target <- sub("^--class=", "", args[[1]])
  } else if (startsWith(mode, "--interface=")) {
    parse_mode <- "interface"
    target <- sub("^--interface=", "", args[[1]])
  } else if (mode %in% c("--class", "-c", "class")) {
    parse_mode <- "class"
    target <- extract_target(args[[1]], args)
  } else if (mode %in% c("--interface", "-i", "interface", "iface")) {
    parse_mode <- "interface"
    target <- extract_target(args[[1]], args)
  } else if (mode %in% c("--all", "-a", "all")) {
    parser$parse(exclude_files = exclude_files)
  } else {
    # Backwards compatibility: single argument treated as interface name
    parse_mode <- "interface"
    target <- args[[1]]
  }

  if (!is.null(parse_mode)) {
    if (!nzchar(target)) {
      cli::cli_abort("Class name cannot be empty.")
    }

    if (parse_mode == "class") {
      parser$parse_class(target, exclude_files = exclude_files)
    } else {
      parser$parse_interface_for_class(target, exclude_files = exclude_files)
    }
  }
}

# Optional extras
parser$detect_usages(exclude_files = exclude_files)
parser$extract_enums(exclude_files = exclude_files)

library(here)

source(here("automate/codes/class-doc.R"))

classes_folder <- here("src")

parser <- VBADocParser$new(
  folder = classes_folder,
  proj_path = here::here(),
  output_folder = here("src", "docs")
)

args <- commandArgs(trailingOnly = TRUE)
exclude_impl <- here("src/classes/implements/BetterArray.cls")

if (length(args) == 0) {
  parser$parse(exclude_files = exclude_impl)
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
    parser$parse(exclude_files = exclude_impl)
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
      parser$parse_class(target, exclude_files = exclude_impl)
    } else {
      parser$parse_interface_for_class(target, exclude_files = exclude_impl)
    }
  }
}

# Optional extras
parser$detect_usages(exclude_files = exclude_impl)
parser$extract_enums(exclude_files = exclude_impl)

# Build master markdown and a simple index page
parser$build_master_markdown(title = "Outbreak Tools – Code Documentation")
parser$build_site_index(title = "Outbreak Tools – Code Documentation")

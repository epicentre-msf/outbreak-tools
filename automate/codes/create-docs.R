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

if (length(args) > 0) {
  # Generate documentation for the interface of a single class passed as first arg
  target_class <- args[[1]]
  parser$parse_interface_for_class(target_class, exclude_files = exclude_impl)
} else {
  # Generate documentation for all classes
  parser$parse(exclude_files = exclude_impl)
}

# Optional extras
parser$detect_usages(exclude_files = exclude_impl)
parser$extract_enums(exclude_files = exclude_impl)

# Build master markdown and a simple index page
parser$build_master_markdown(title = "Outbreak Tools – Code Documentation")
parser$build_site_index(title = "Outbreak Tools – Code Documentation")

library(here)

source(here("automate/codes/class-doc.R"))

classes_folder <- here("src/classes")

parser <- VBADocParser$new(
  folder = classes_folder,
  proj_path = here::here(),
  output_folder = here("docs")
)

parser$parse(exclude_files = here("src/classes/implements/BetterArray.cls"))
parser$extract_enums(
  exclude_files = here("src/classes/implements/BetterArray.cls")
)

quarto::quarto_render("docs/CustomTable.md")

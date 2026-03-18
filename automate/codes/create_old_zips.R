# One-time script to create missing zip files in releases/old/
# Scans releases/main/ and releases/dev/ for complete release sets
# (designer + setup + ribbon) and creates zips for dates not already in old/

pacman::p_load(fs, glue, here)

old_folder <- here("releases", "old")
dir_create(old_folder)

# Get existing zips in old/
existing_zips <- dir_ls(old_folder, glob = "*.zip") |>
  path_file()

create_zip_for_branch <- function(branch) {
  designer_dir <- here("releases", branch, "designer")
  setup_dir <- here("releases", branch, "setup")
  ribbon_dir <- here("releases", branch, "ribbon")

  if (!dir_exists(designer_dir) || !dir_exists(setup_dir) || !dir_exists(ribbon_dir)) {
    message(glue("Skipping {branch}: missing designer/setup/ribbon folders"))
    return(invisible(NULL))
  }

  # List all designer files and extract dates
  designers <- dir_ls(designer_dir, glob = "*.xlsb") |> path_file()
  # Extract date pattern YYYY-MM-DD from filenames
  dates <- regmatches(designers, regexpr("\\d{4}-\\d{2}-\\d{2}", designers))

  for (rel_date in dates) {
    zip_name <- glue("OBT-{branch}-{rel_date}.zip")

    if (zip_name %in% existing_zips) {
      message(glue("  Already exists: {zip_name}"))
      next
    }

    # Find matching files for this date
    designer_file <- dir_ls(designer_dir, glob = glue("*{rel_date}*"))
    setup_file <- dir_ls(setup_dir, glob = glue("*{rel_date}*"))
    ribbon_file <- dir_ls(ribbon_dir, glob = glue("*{rel_date}*"))

    # Skip if we don't have all 3 components
    if (length(designer_file) == 0 || length(setup_file) == 0 || length(ribbon_file) == 0) {
      message(glue("  Skipping {rel_date} ({branch}): incomplete set"))
      next
    }

    # Create zip with flat structure (no directory paths inside)
    zip_path <- here(old_folder, zip_name)
    utils::zip(
      zipfile = zip_path,
      files = c(designer_file[1], setup_file[1], ribbon_file[1]),
      flags = "-j"
    )
    message(glue("  Created: {zip_name}"))
  }
}

message("Creating missing old zips for main branch...")
create_zip_for_branch("main")

message("\nCreating missing old zips for dev branch...")
create_zip_for_branch("dev")

message("\nDone. Contents of releases/old/:")
dir_ls(old_folder) |> path_file() |> sort() |> cat(sep = "\n")

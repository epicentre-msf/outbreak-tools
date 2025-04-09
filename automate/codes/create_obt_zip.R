# copy the lastest designer and setup files

# loading required libraries
pacman::p_load(
  fs,
  glue,
  lubridate,
  dplyr,
  here
)


branch <- system("git branch --show-current", intern = TRUE)

# the actual branch is infered from the git branch you are currently ons
branch <- dplyr::case_when(
  branch == "dev" ~ "dev",
  branch == "hot-fixes" ~ "dev",
  branch == "main" ~ "main"
)


# the current outbreak tools folder is infered from the environment variable
# created on your system. Make sure to set it up either in the project .Renviron
# file or in your system environment variables.
obt_folder <- Sys.getenv("OBTFOLDER")


create_obt_release <- function(actual_branch = branch, latest = FALSE) {
  # rel_date is the release date in format yyyy-mm-dd
  rel_date <- today()

  # copy files to the release folder, create it if it does not exist
  rel_folder <- here("releases", glue("{actual_branch}")) |> dir_create()

  # copy the actual designer file to the release folder
  file_copy(
    path = here("src", "bin", "designer", "designer.xlsb"),
    new_path = here(
      glue("{rel_folder}"),
      "designer",
      glue("designer_{actual_branch}-{rel_date}.xlsb")
    ),
    overwrite = TRUE
  )

  # copy the ribbon template to the release folder
  file_copy(
    path = here("ribbons", glue("_ribbontemplate_{actual_branch}.xlsb")),
    new_path = here(
      glue("{rel_folder}"),
      "ribbon",
      glue("_ribbontemplate_{actual_branch}-{rel_date}.xlsb")
    ),
    overwrite = TRUE
  )

  # copy the setup to the release folder
  file_copy(
    path = here("src", "bin", "setup", "setup.xlsb"),
    new_path = here(
      glue("{rel_folder}"),
      "setup",
      glue("setup_{actual_branch}-{rel_date}.xlsb")
    ),
    overwrite = TRUE
  )

  # if latest, add the zip file and update the latest folder.
  if (latest) {
    # create the latest folder if it does not exist

    latest_folder <- here(
      "release",
      "latest",
      glue("OBT-{actual_branch}-latest")
    ) |>
      dir_create()

    # remove all the files in the latest OBT folder (the folder should be empty)
    file_delete(
      path = here(
        "release",
        "latest",
        glue("OBT-{actual_branch}-latest")
      ) |>
        dir_ls()
    )

    # copy the designer, the setup and the ribbon template
    # from the release folder to the latest folder
    file_copy(
      path = here(
        "releases",
        glue("{actual_branch}"),
        "designer",
        glue("designer_{actual_branch}-{rel_date}.xlsb")
      ),
      new_path = here(
        glue("{latest_folder}"),
        glue("designer_{actual_branch}-{rel_date}.xlsb")
      )
    )

    file_copy(
      path = here(
        "releases",
        glue("{actual_branch}"),
        "setup",
        glue("setup_{actual_branch}-{rel_date}.xlsb")
      ),
      new_path = here(
        glue("{latest_folder}"),
        glue("setup_{actual_branch}-{rel_date}.xlsb")
      )
    )

    file_copy(
      path = here(
        "releases",
        glue("{actual_branch}"),
        "ribbon",
        glue("_ribbontemplate_{actual_branch}-{rel_date}.xlsb")
      ),
      new_path = here(
        glue("{latest_folder}"),
        glue("_ribbontemplate_{actual_branch}-{rel_date}.xlsb")
      )
    )

    # add the files to a zip file for releases. The first zip file is the file
    # with the corresponding release date.
    utils::zip(
      zipfile = here(
        "releases",
        "old",
        glue("OBT-{actual_branch}-{rel_date}.zip")
      ),
      files = c(
        here(
          glue("{latest_folder}"),
          glue("setup_{actual_branch}-{rel_date}.xlsb")
        ),
        here(
          glue("{latest_folder}"),
          glue("designer_{actual_branch}-{rel_date}.xlsb")
        ),
        here(
          glue("{latest_folder}"),
          glue("_ribbontemplate_{actual_branch}-{rel_date}.xlsb")
        )
      ),
      flags = "-j"
    )

    # copy the new release to the old folder
    file_copy(
      here("releases", "old", glue("OBT-{actual_branch}-{rel_date}.zip")),
      here("latest", glue("OBT-{actual_branch}-latest.zip")),
      overwrite = TRUE
    )
  }
}

# create the release for the master setup
create_master_setup_release <- function(
  actual_branch = branch,
  latest = FALSE
) {
  rel_date <- today()
  # copy files to the release folder, create it if it does not exist
  rel_folder <- here("releases", glue("{actual_branch}")) |> dir_create()
  # copy the master setup file to the corresponding folder
  file_copy(
    here("src", "bin", "master-setup", "master_setup.xlsb"),
    here(
      glue("{rel_folder}"),
      "master-setup",
      glue("master_setup_{actual_branch}-{rel_date}.xlsb")
    ),
    overwrite = TRUE
  )

  if (latest) {
    # add the file to the lastest folder
    file_copy(
      here("src", "bin", "master-setup", "master_setup.xlsb"),
      here("releases", "latest", "master_setup-latest.xlsb")
    )

    # copy to old versions
    file_copy(
      here("src", "bin", "master-setup", "master_setup.xlsb"),
      here("releases", "old", glue("master_setup-{rel_date}.xlsb")),
      overwrite = TRUE
    )
  }
}

# create_obt_release(latest = TRUE)
# create_master_setup_release(latest = TRUE)

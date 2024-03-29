# copy the lastest designer and setup files
user_name  <- Sys.info()[["user"]]
actual_branch <- system("git branch --show-current", intern = TRUE)

# Please pay attention to which branch you are copying from.
# SHOULD BE ON DEV BRANCH

obt_folder  <- dplyr::case_when(
    user_name == "Y-AMEVOIN" ~ "D:/Projects/outbreak-tools",
    user_name == "komlaviamevoin" ~ "~/Documents/Projects/outbreak-tools",
    TRUE ~ "~/outbreak-tools"
)

setup_folder  <- dplyr::case_when(
    user_name == "Y-AMEVOIN" ~ "D:/Projects/outbreak-tools-setup",
    user_name == "komlaviamevoin" ~ "~/Documents/Projects/outbreak-tools-setup",
    TRUE ~ "~/outbreak-tools-setup"
)

# copy the designer
rel_date <- lubridate::today()
obt_all_folder <- glue::glue("./src/OBT_all_{actual_branch}")
if (!dir.exists(obt_all_folder)) dir.create(obt_all_folder)

file.copy(from = glue::glue("{obt_folder}/src/bin/designer_{actual_branch}.xlsb"),
          to = glue::glue("{obt_all_folder}/",
                          "designer_{actual_branch}-{rel_date}.xlsb"),
          overwrite = TRUE)

# # copy the ribbon template
file.copy(
    from = glue::glue("{obt_folder}/misc/",
                      "_ribbontemplate_{actual_branch}.xlsb"),
    to = glue::glue("{obt_all_folder}/",
                    "_ribbontemplate_{actual_branch}-{rel_date}.xlsb"),
    overwrite = TRUE
)



# # copy the empty setup
file.copy(
    from = glue::glue("{setup_folder}/setup.xlsb"),
    to = glue::glue("{obt_all_folder}/setup-{rel_date}.xlsb"),
    overwrite = TRUE
)

file.remove(glue::glue("src/OBT_all_{actual_branch}.zip"))
# add the files to a zip file for demo
utils::zip(
    zipfile = glue::glue("src/OBT_all_{actual_branch}.zip"),
    files = c(
    glue::glue("{obt_all_folder}/setup-{rel_date}.xlsb"),
    glue::glue("{obt_all_folder}/designer_{actual_branch}-{rel_date}.xlsb"),
    glue::glue("{obt_all_folder}/_ribbontemplate_{actual_branch}-{rel_date}.xlsb"),
    "./automation/run_designer_on_windows.R",
    "./automation/rundesigner.vbs"
    ),
    flags = "-j"
)

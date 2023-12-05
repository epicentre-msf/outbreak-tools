# copy the lastest designer and setup files
user_name  <- Sys.info()[["user"]]

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

file.copy(from = glue::glue("{obt_folder}/designer.xlsb"),
          to = glue::glue("./src/OBT_all/designer-{rel_date}.xlsb"),
          overwrite = TRUE)

# # copy the ribbon template
file.copy(from = glue::glue("{obt_folder}/misc/_ribbontemplate.xlsb"),
          to = "./src/OBT_all/_ribbontemplate.xlsb",
          overwrite = TRUE
)

# # copy the empty setup
file.copy(
    from = glue::glue("{setup_folder}/setup.xlsb"),
    to = glue::glue("./src/OBT_all/setup-{rel_date}.xlsb"),
    overwrite = TRUE
)

file.remove("src/OBT_all.zip")
# add the files to a zip file for demo
utils::zip(
    zipfile = "src/OBT_all.zip",
    files = c(
    glue::glue("./src/OBT_all/setup-{rel_date}.xlsb"),
    glue::glue("./src/OBT_all/designer-{rel_date}.xlsb"),
    glue::glue("./src/OBT_all/_ribbontemplate.xlsb"),
    "./automation/run_designer_on_windows.R",
    "./automation/rundesigner.vbs"
    ),
    flags = "-j"
)

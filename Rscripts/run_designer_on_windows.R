# This is a quick R script to run the designer from R will only work on
# windows OS.
# I used here, and glue packages with R version > 4.1 (pipe operator)

used_packages <- c("here", "glue")
to_install <- setdiff(used_packages, installed.packages()[, 1])
if (length(to_install)) install.packages(used_packages)

# Add relative paths
outbreak_tools_path <- here::here()

# a simple function to set the path

set_path <- function(file_name, out_path = outbreak_tools_path) {
  glue::glue("{out_path}/{file_name}") |>
    normalizePath() |>
    shQuote()
}

# Define the parameters


designer_path <- set_path("linelist_designer_aky.xlsb")
setup_path <- set_path("input/outbreak-tools-setup/setup.xlsb")
#The geobase is optional
geo_path <- set_path("input/geobase/geobase_obt_yem_20230112.xlsx")
output_dir <- set_path("output")
linelist_name <- shQuote("rinterface_test") # Name of the linelist file
setup_lang <- shQuote("English") # Language of the dictionary
linelist_lang <- shQuote("English") # Language of the linelist interface

# Sending code to the designer

cmd <- glue::glue("{outbreak_tools_path}/Rscripts/rundesigner.vbs",
                  " {designer_path} {geo_path} {setup_path}",
                  " {output_dir} {linelist_name}",
                  " {setup_lang} {linelist_lang}")


# run the shell command

shell(cmd, "cscript", flag = "//nologo")

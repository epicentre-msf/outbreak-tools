# RUN THE DESIGNER FROM R
# Yves Amevoin, Epicentre
# Last modification date: 10 November 2023
 
# Description:
# ------------------------------------------------------------------------------
# This is a quick R script to run the designer from R will only work on Windows OS.
# Requires R version > 4.1 (pipe operator) and packages: glue

# Requirements:
# ------------------------------------------------------------------------------
# R version > 4.1
# ONLY WORKS ON WINDOWS OS
# SET YOUR WORKING DIRECTORY TO THE SOURCE FILE LOCATION
# Make sure you have the *.vbs (rundesigner.vbs) file in the working directory

# INITIALISATIONS
# ------------------------------------------------------------------------------
#
# install required packages
used_packages <- "glue"
to_install <- setdiff(used_packages, installed.packages()[, 1])
if (length(to_install)) install.packages(used_packages)


# function to define correctly the paths for shell command
obt_folder <- getwd()

gl <- glue::glue

set_path <- function(file_name, path = obt_folder) {
  glue::glue("{path}/{file_name}") |>
    normalizePath() |>
    shQuote()
}

# DEFINE THE PARAMETERS REQUIRED BY THE DESIGNER
# ------------------------------------------------------------------------------

# Output directory for the linelist (in the working directory)
# shQuote is for adding quotes for the shell commands
output_dir <- "output_linelist"

# create the output directory if it does not exists
if (!dir.exists(gl("{obt_folder}/output_linelist"))){
  dir.create(gl("{obt_folder}/output_linelist"))
}

# test if the rundesigner vbs file is in the current working directory.
if (!file.exists(gl("{obt_folder}/rundesigner.vbs"))){
  warning(gl("Unable to find the .vbs file in the folder {obt_folder}"))
}

# precise the parameters (you can change the name of the files)

designer_path <- set_path("designer_aky.xlsb") # Designer_path
setup_path <- set_path("setup.xlsb") # Setup path
ribbon_path <- set_path("_ribbontemplate.xlsb") # Path to the template ribbon file
output_dir <- set_path("output_linelist")
setup_lang <- shQuote("English") # Language of the dictionary
linelist_lang <- shQuote("English") # Language of the Linelist interface


#Name of the linelist file (you can modify the name)
linelist_name <- shQuote("rinterface_test") 

# The geobase path is optional, you can ignore
geo_path <- shQuote("")

# Sending code to the designer
cmd <- gl(
  "{obt_folder}/rundesigner.vbs",
  " {designer_path} {geo_path} {setup_path}",
  " {output_dir} {linelist_name}",
  " {setup_lang} {linelist_lang} {ribbon_path}"
)

# This shell command runs the vbs script behind and generate the linelist
# using the parameters you provide. The vbs script is the rundesigner.vbs file
# located in your working directory
shell(cmd, "cscript", flag = "//nologo")

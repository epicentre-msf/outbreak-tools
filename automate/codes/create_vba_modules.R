pacman::p_load(here, glue)

module_header <- function(
  module_name,
  module_description = "",
  folder_name = ""
) {
  outfilename = glue::glue("{module_name}.bas")

  file_content <- glue::glue(
    "
Attribute VB_Name = \"{module_name}\"

Option Explicit

'@Folder(\"{foldername}\")
'@IgnoreModule UnrecognizedAnnotation, SheetAccessedUsingString, ParameterCanBeByVal, ParameterNotUsed : some parameters of controls are not used
    "
  )
}

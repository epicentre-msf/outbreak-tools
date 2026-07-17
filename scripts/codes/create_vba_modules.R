pacman::p_load(here, glue, xfun)

# @section Header
# ===============================================================================
module_header <- function(
  module_name,
  folder_name = "Modules",
  module_description = ""
) {
  glue::glue(
    "
Attribute VB_Name = \"{module_name}\"

Option Explicit

'@Folder(\"{folder_name}\")
'@ModuleDescription(\"{module_description}\")
'@IgnoreModule UnrecognizedAnnotation, SheetAccessedUsingString, ParameterCanBeByVal, ParameterNotUsed
"
  )
}

write_crlf <- function(path, contents) {
  normalised <- gsub("\r?\n", "\n", contents, perl = TRUE)
  crlf_text <- gsub("\n", "\r\n", normalised, fixed = TRUE)
  con <- file(path, open = "wb")
  on.exit(close(con))
  writeBin(charToRaw(crlf_text), con)
}

# Create an empty .bas module under src/modules/<folder>/.
# `module_path` may be a bare name ("MyModule") or a topic path
# ("topic/MyModule"); the topic becomes both the destination subfolder and the
# @Folder annotation.
create_modules <- function(
  module_path,
  module_description = ""
) {
  module_name <- module_path |> basename() |> xfun::sans_ext()
  folder_name <- module_path |> dirname() |> basename()
  if (folder_name %in% c(".", "", "/")) {
    folder_name <- "Modules"
    out_dir <- here("src", "modules")
  } else {
    out_dir <- here("src", "modules", folder_name)
  }

  dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
  out_path <- file.path(out_dir, glue("{module_name}.bas"))
  write_crlf(out_path, module_header(module_name, folder_name, module_description))
  message(glue("Created module: {out_path}"))
}

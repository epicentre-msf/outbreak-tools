# create en empty file for creating an interface for each of the class
pacman::p_load(here, glue)

# @section Headers
# ===============================================================================
class_header <- function(
  class__name,
  module_description = "",
  interface = FALSE
) {
  # nolint
  interf <- ifelse(interface, "I", "")
  predeclare_id <- ifelse(interface, "False", "True")

  glue::glue(
    "
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = \"{interf}{class_name}\"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = {predeclare_id}
Attribute VB_Exposed = False
Attribute VB_Description = \"{module_description}\"

'@Folder(\"Dictionary\")
'@ModuleDescription(\"{module_description}\")
'@IgnoreModule UnrecognizedAnnotation

Option Explicit
"
  )
}

test_header <- function(class_name) {
  glue::glue(
    "
Attribute VB_Name = \"Test{class_name}\"
Attribute VB_Description = \"Tests for {class_name} class \"

Option Explicit
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder(\"CustomTests\")
'@ModuleDescription(\"Tests for {class_name} class \")

Private Const TEST_OUTPUT_SHEET As String = \"testsOutputs\"
'add other private constants if required
  
Private Assert As ICustomTest

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    'this method runs once per module.
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName = \"Test{class_name}\"

    'Add other importants initialization logic
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'This method runs before every test in the module..

'@TestInitialize
Private Sub TestInitialize()

End Sub
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


create_class <- function(
  class_path,
  description = ""
) {
  class_name <- class_path |> basename() |> xfun::sans_ext()
  folder_name <- class_path |> dirname() |> basename()
  # nolint

  class_name_header <- class_header(
    class_name,
    description = description,
    module_description = module_description
  ) # nolint
  class_interface_header <- class_header(
    class_name,
    interface = TRUE,
    description = description,
    module_description = glue::glue("Interface of {module_description}")
  ) # nolint

  class_test_header <- test_header(class_name)
  # create the class
  write_crlf(
    here("src", "classes", "implements", glue("{class_name}.cls")),
    class_name_header
  )
  # create the interface of the class
  write_crlf(
    here("src", "classes", "interfaces", glue("I{class_name}.cls")),
    class_interface_header
  )
  # Add Test for the class
  write_crlf(
    here("src", "tests", glue("Test{class_name}.bas")),
    class_test_header
  )
}

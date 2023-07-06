# create en empty file for creating an interface for each of the class

class_header  <- function(class_name, description = "", module_description = "",
 interface = FALSE){ #nolint
interf  <- ifelse(interface, "I", "")
predeclare_id  <- ifelse(interface, "False", "True")

glue::glue("
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = \"{interf}{class_name}\"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = {predeclare_id}
Attribute VB_Exposed = False
Attribute VB_Description = \"{description}\"

'@Folder(\"Dictionary\")
'@ModuleDescription(\"{module_description}\")
'@IgnoreModule

Option Explicit

'Exposed methods
")

}

test_header  <- function(class_name) {

glue::glue(
"
Attribute VB_Name = \"Test{class_name}\"

Option Explicit
Option Private Module

'@TestModule
'@Folder(\"Tests\")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject(\"Rubberduck.AssertClass\")
    Set Fakes = CreateObject(\"Rubberduck.FakesProvider\")
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
")
}

create_class  <- function(class_name, description = "", module_description = ""){ #nolint


        class_name_header  <- class_header(class_name,
                                           description = description,
                                           module_description = module_description) #nolint
        class_interface_header  <- class_header(class_name, interface = TRUE,
                                                description = description,
                                                module_description = glue::glue("Interface of {module_description}")) #nolint

        class_test_header  <- test_header(class_name)
        #create the class
        cat(class_name_header,
            file = glue::glue("./src/classes/implements/{class_name}.cls"))
        #create the interface of the class
        cat(class_interface_header,
        file = glue::glue("./src/classes/interfaces/I{class_name}.cls"))
        # Add Test for the class
        cat(class_test_header,
        file = glue::glue("./src/modules/tests/Test{class_name}.bas"))
}

# You can create the class here by precising the class name
#create_class("UpVal") #nolint
